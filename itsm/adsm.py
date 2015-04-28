#!/usr/bin/env python
# coding=utf-8

import collections

from nltk import metrics, stem, tokenize
from suds.sax.element import Attribute, Element

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
		cache_key = '%s/%s' % (list_uuid, ','.join(viewFields))
		table = self.listitem_ref._cache.get(cache_key)
		if not table:
			if not query:
				query = Element('ns1:query').append(Element('Query').append(Element('Where').append(Element('IsNotNull').append(Element('FieldRef').append(Attribute('Name', 'ID'))))))
			if viewFields:
				fields = Element('ViewFields')
				for f in viewFields:
					fields.append(Element('FieldRef').append(Attribute('Name', f)))
				fields = Element('ns1:viewFields').append(fields)
			else:
				fields = None
			list_items = self.adsm_lists.service.GetListItems(list_uuid, query=query, viewFields=fields, rowLimit=9999)
			list_items_rows = list_items.listitems.data.row if int(list_items.listitems.data._ItemCount) > 1 \
			          else [list_items.listitems.data.row] if int(list_items.listitems.data._ItemCount) > 0 \
			          else []
			table = dict(filter(lambda x: x[0], map(lambda x: (x.__dict__.get(field), x), list_items_rows)))
			self.listitem_ref._cache[cache_key] = table

		# Attempt to get exact match
		match = table.get(field_value)
		if match:
			return '%s%s%s' % (match['_ows_ID'], ADSMBase.ref_sep, match[display_field] if display_field else '')

		# If fuzzy is True, attempt to find a reasonable match
		if fuzzy:
			stemmer = stem.PorterStemmer()
			def normalize(s):
				words = tokenize.wordpunct_tokenize(s.lower().strip())
				return ' '.join([stemmer.stem(w) for w in words])

			normalized_cache_key = '%s/normalized' % cache_key
			normalized_table = self.listitem_ref._cache.get(normalized_cache_key)
			if not normalized_table:
				normalized_table = dict((normalize(k), v) for k, v in table.items())
				self.listitem_ref._cache[normalized_cache_key] = normalized_table

			normalized_field_value = normalize(field_value)
			candidates = sorted(normalized_table.items(), lambda x, y: metrics.edit_distance(x, normalized_field_value) - metrics.edit_distance(y, normalized_field_value), lambda t: t[0])
			if metrics.edit_distance(candidates[0][0], normalized_field_value) <= max_dist:
				match = candidates[0][1]
				return '%s%s%s' % (match['_ows_ID'], ADSMBase.ref_sep, match[display_field] if display_field else '')

		# No match found
		return None
	listitem_ref._cache = {}

	def listitem_refs(self, list_uuid, query, viewFields, field, field_values, display_field='_ows_Title', fuzzy=False, max_dist=4):
		if not field_values:
			return None
		if isinstance(field_values, basestring):
			field_values = [field_values]
		return ADSMBase.ref_sep.join(filter(lambda x: x, map(lambda field_value: self.listitem_ref(list_uuid, query, viewFields, field, field_value, display_field=display_field, fuzzy=fuzzy, max_dist=max_dist), field_values)))

	def reset_caches(self):
		self.listitem_ref._cache = {}
		self.person_ref._cache = {}

	# Sync functions

	def sync_to_list_by_comparison(self, list_uuid, query, viewFields, list_items_compare_key, ext_items, ext_items_compare_key, compare_f, field_map, content_type='Item', folder=None, fuzzy=False, max_dist=4, commit=True):
		if not query:
			query = Element('ns1:query').append(Element('Query').append(Element('Where').append(Element('IsNotNull').append(Element('FieldRef').append(Attribute('Name', 'ID'))))))
		if viewFields:
			fields = Element('ViewFields')
			for f in viewFields:
				fields.append(Element('FieldRef').append(Attribute('Name', f)))
			fields = Element('ns1:viewFields').append(fields)
		else:
			fields = None
		list_items = self.adsm_lists.service.GetListItems(list_uuid, query=query, viewFields=fields, rowLimit=9999)
		list_items_rows = list_items.listitems.data.row if int(list_items.listitems.data._ItemCount) > 1 \
		            else [list_items.listitems.data.row] if int(list_items.listitems.data._ItemCount) > 0 \
		            else []
		list_items_map = dict(filter(lambda p: p[0], map(lambda x: (x.__dict__.get(list_items_compare_key), x), list_items_rows)))
		if fuzzy:
			stemmer = stem.PorterStemmer()
			def normalize(s):
				words = tokenize.wordpunct_tokenize(s.lower().strip())
				return ' '.join([stemmer.stem(w) for w in words])
			list_items_normalized_map = dict((normalize(k), v) for k, v in list_items_map.items())
		
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
			list_item = list_items_map.get(ext_item[ext_items_compare_key])
			if not list_item and fuzzy:
				normalized_compare_v = normalize(ext_item[ext_items_compare_key])
				candidates = sorted(list_items_normalized_map.items(), lambda x, y: metrics.edit_distance(x, normalized_compare_v) - metrics.edit_distance(y, normalized_compare_v), lambda t: t[0])
				if metrics.edit_distance(candidates[0][0], normalized_compare_v) <= max_dist:
					list_item = candidates[0][1]
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

