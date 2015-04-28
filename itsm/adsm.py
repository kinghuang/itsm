#!/usr/bin/env python
# coding=utf-8

import collections
import os

from nltk import metrics, stem, tokenize
from suds.sax.element import Attribute, Element
from suds.sax.parser import Parser

from itsm.base import Base


class ADSMBase(Base):

	ref_sep = ';#'
	batch_size = 20

	def argument_parser(self):
		parser = super(ADSMBase, self).argument_parser()

		parser.add_argument('env', default='SP_ADSM', nargs='?')
		parser.add_argument('-d', help='dry run', action='store_true')

		return parser

	def _delayed_adsm_client(self, var, sp):
		if not hasattr(self, var):
			setattr(self, var, self._named_client(self.args.env, lambda url, **x: self.create_sharepoint_client(url + '/' + sp, **x)))
		return getattr(self, var)

	# Access web services clients for ADSM endpoints

	@property
	def adsm_lists(self):     return self._delayed_adsm_client('_adsm_lists',     'Lists.asmx?WSDL')
	@property
	def adsm_people(self):    return self._delayed_adsm_client('_adsm_people',    'People.asmx?WSDL')
	@property
	def adsm_webs(self):      return self._delayed_adsm_client('_adsm_webs',      'Webs.asmx?WSDL')
	@property
	def adsm_usergroup(self): return self._delayed_adsm_client('_adsm_usergroup', 'UserGroup.asmx?WSDL')

	@property
	def ci_list_uuid(self): return os.environ[self.args.env + '_CI_LIST']
	@property
	def uc_list_uuid(self): return os.environ[self.args.env + '_UC_LIST']
	@property
	def choices_list_uuid(self): return os.environ[self.args.env + '_CHOICES_LIST']

	# Caching support
	def _cachetable(self, name):
		if not hasattr(self, '_cachetables'):
			setattr(self, '_cachetables', {})
		tables = getattr(self, '_cachetables')

		table = tables.get(name)
		if not table:
			table = {}
			tables[name] = table

		return table

	def reset_caches(self):
		setattr(self, '_cachetables', {})

	# List functions

	def listitems(self, list_uuid, query=None, fields=('ID', 'Title'), limit=9999, cache=True):
		if cache:
			table = self._cachetable('_get_listitems')
			key = '%s/%s/%s/%s' % (list_uuid, query, fields, limit)
			if key in table:
				return table[key]

		query = Element('ns1:query').append(Element('Query').append(Element('Where').append(Element('IsNotNull').append(Element('FieldRef').append(Attribute('Name', 'ID')))))) if not query else query
		fields = Element('ns1:viewFields').append(Element('ViewFields').append([Element('FieldRef').append(Attribute('Name', f)) for f in fields])) if fields else None

		list_items = self.adsm_lists.service.GetListItems(list_uuid, query=query, viewFields=fields, rowLimit=limit)
		list_items_rows = list_items.listitems.data.row if int(list_items.listitems.data._ItemCount) > 1 \
			          else [list_items.listitems.data.row] if int(list_items.listitems.data._ItemCount) > 0 \
			          else []
		if cache:
			table[key] = list_items_rows

		return list_items_rows

	def keyed_listitems(self, list_items, key_field='_ows_Title'):
		return dict(filter(lambda x: x[0], map(lambda x: (x.__dict__.get(key_field), x), list_items)))

	def fuzzy_keyed_listitems(self, list_items, key_field='_ows_Title'):
		return dict(filter(lambda x: x[0], map(lambda x: (self.normalize(x.__dict__.get(key_field)), x), list_items)))

	def normalize(self, s, stemmer=stem.PorterStemmer()):
		words = tokenize.wordpunct_tokenize(s.lower().strip())
		return ' '.join([stemmer.stem(w) for w in words])

	def fuzzy_match(self, fuzzy_keyed_list_items, s, max_dist=4):
		normalized_key = self.normalize(s)
		candidates = filter(lambda y: y[0] <= max_dist, ((metrics.edit_distance(x[0], normalized_key), x[1]) for x in fuzzy_keyed_list_items.items()))
		candidates.sort(lambda x, y: x[0] - y[0])

		return candidates[0][1] if len(candidates) > 0 else None

	# Reference functions

	def person_ref(self, principal):
		if not principal:
			return None
		if principal in self.person_ref._cache:
			return self.person_ref._cache.get(principal)

		def resolve(p):
			principals = Element('ns1:principalKeys').append(Element('ns1:string').setText(p))
			result = self.adsm_people.service.ResolvePrincipals(principals, 'User', True)
			return result.PrincipalInfo[0]

		candidate = resolve(principal)
		if not candidate.IsResolved:
			self.person_ref._cache[principal] = None
			return None
		elif candidate.AccountName.startswith('UC_ADMIN') or candidate.AccountName.startswith('UC_CAMPUS'):
			# if the resolved candidate is a UC_ADMIN or UC_CAMPUS account,
			# attempt to find an equivalent UC account
			uc_principal = candidate.AccountName.replace('UC_ADMIN', 'UC').replace('UC_CAMPUS', 'UC')
			uc_candidate = resolve(uc_principal)
			if uc_candidate.IsResolved:
				candidate = uc_candidate

		ref = '%s%s%s' % (candidate.UserInfoID, ADSMBase.ref_sep, candidate.DisplayName)
		self.person_ref._cache[principal] = ref
		return ref
	person_ref._cache = {}

	def people_refs(self, principals):
		if not principals:
			return None
		if isinstance(principals, basestring):
			principals = [principals]
		return ADSMBase.ref_sep.join(filter(lambda x: x, map(self.person_ref, principals)))

	def listitem_ref(self, list_uuid, query, viewFields, field, field_value, display_field='_ows_Title', fuzzy=False, max_dist=4):
		cache = self._cachetable('listitems')
		cache_key = '%s/%s/%s/%s' % (list_uuid, query, viewFields, field)

		table = cache.get(cache_key)
		if not table:
			listitems = self.listitems(list_uuid, query=query, fields=viewFields)
			table = self.keyed_listitems(listitems, key_field=field)
			cache[cache_key] = table

		# Attempt to get exact match
		match = table.get(field_value)
		
		# If no exact match is found and fuzzy is True, attempt to find a reasonable match
		if not match and fuzzy:
			fuzzy_cache_key = '%s/fuzzy' % cache_key

			fuzzy_table = cache.get(fuzzy_cache_key)
			if not fuzzy_table:
				listitems = self.listitems(list_uuid, query=query, fields=viewFields)
				fuzzy_table = self.fuzzy_keyed_listitems(listitems, key_field=field)
				cache[fuzzy_cache_key] = fuzzy_table

			match = self.fuzzy_match(fuzzy_table, field_value, max_dist=max_dist)

		return '%s%s%s' % (match['_ows_ID'], ADSMBase.ref_sep, match[display_field] if display_field else '') if match else None

	def listitem_refs(self, list_uuid, query, viewFields, field, field_values, display_field='_ows_Title', fuzzy=False, max_dist=4):
		if not field_values:
			return None
		if isinstance(field_values, basestring):
			field_values = [field_values]
		return ADSMBase.ref_sep.join(filter(lambda x: x, map(lambda field_value: self.listitem_ref(list_uuid, query, viewFields, field, field_value, display_field=display_field, fuzzy=fuzzy, max_dist=max_dist), field_values)))

	def ci_ref(self, content_type, field, field_value, display_field='_ows_Title', fuzzy=False, max_dist=4):
		return self._known_listitem_ref('ci', content_type, field, field_value, display_field=display_field, fuzzy=fuzzy, max_dist=max_dist)

	def uc_ref(self, content_type, field, field_value, display_field='_ows_Title', fuzzy=False, max_dist=4):
		return self._known_listitem_ref('uc', content_type, field, field_value, display_field=display_field, fuzzy=fuzzy, max_dist=max_dist)

	def choice_ref(self, content_type, field, field_value, display_field='_ows_Title', fuzzy=False, max_dist=4):
		return self._known_listitem_ref('choices', content_type, field, field_value, display_field=display_field, fuzzy=fuzzy, max_dist=max_dist)

	def _known_listitem_ref(self, listname, content_type, field, field_value, display_field='_ows_Title', fuzzy=False, max_dist=4):
		list_uuid = getattr(self, '%s_list_uuid' % listname)
		query = Parser().parse(string=str('<ns1:query><Query><Where><Eq><FieldRef Name="ContentType"/><Value Type="Text">%s</Value></Eq></Where></Query></ns1:query>' % content_type)).getChild('query') \
		        if content_type else None
		viewFields = ('ID', 'Title', field[5:] if field.startswith('_ows_') else field)

		return self.listitem_ref(list_uuid, query, viewFields, field, field_value, display_field=display_field, fuzzy=fuzzy, max_dist=max_dist)

	# Sync functions

	def sync_to_list_by_comparison(self, list_uuid, query, viewFields, list_items_compare_key, ext_items, ext_items_compare_key, compare_f, field_map, content_type='Item', folder=None, fuzzy=False, max_dist=4, commit=True):
		cache = self._cachetable('listitems')
		cache_key = '%s/%s/%s/%s' % (list_uuid, query, viewFields, field)

		table = cache.get(cache_key)
		if not table:
			listitems = self.listitems(list_uuid, query=query, fields=viewFields)
			table = self.keyed_listitems(listitems, key_field=field)
			cache[cache_key] = table

		if fuzzy:
			fuzzy_cache_key = '%s/fuzzy' % cache_key

			fuzzy_table = cache.get(fuzzy_cache_key)
			if not fuzzy_table:
				listitems = self.listitems(list_uuid, query=query, fields=viewFields)
				fuzzy_table = self.fuzzy_keyed_listitems(listitems, key_field=field)
				cache[fuzzy_cache_key] = fuzzy_table
		
		method_idx = 1
		batch = Element('Batch')\
		       .append(Attribute('OnError', 'Continue'))\
		       .append(Attribute('ListVersion', 1))
		if folder:
			batch.append(Attribute('RootFolder', folder))

		def update(b):
			if commit:
				updates = Element('ns1:updates').append(b)
				print self.adsm_lists.service.UpdateListItems(listName=list_uuid, updates=updates)

		for ext_item in ext_items:
			list_item = table.get(ext_item[ext_items_compare_key])
			if not list_item and fuzzy:
				list_item = self.fuzzy_match(fuzzy_table, ext_item[ext_items_compare_key], max_dist=4)
			
			method_cmd = compare_f(ext_item, list_item)
			if not method_cmd:
				continue
			item_id = list_item['_ows_ID'] if list_item else 'New'

			# Prepare a method for this new or update item
			method = Element('Method')\
			        .append(Attribute('ID', method_idx))\
			        .append(Attribute('Cmd', method_cmd))\
			        .append(Element('Field')\
			          .append(Attribute('Name', 'ID'))\
			          .setText(item_id))\
			        .append(Element('Field')\
			        	.append(Attribute('Name', 'ContentType'))\
			        	.setText(content_type))
			
			for dst, src in field_map:
				try:
					if not isinstance(src, basestring):
						v = src(ext_item)
					elif isinstance(ext_item, collections.Mapping):
						v = ext_item.get(src)
					else:
						v = getattr(ext_item, src, None)
				except:
					v = None
				e = Element('Field')\
				   .append(Attribute('Name', dst))\
				   .setText(v)
				method.append(e)

			batch.append(method)
			print method
			method_idx += 1

			if len(batch) > ADSMBase.batch_size:
				update(batch)
				batch.detachChildren()

		if len(batch) > 0:
			update(batch)

