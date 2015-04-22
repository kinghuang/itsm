#!/usr/bin/env python
# coding=utf-8

from suds.sax.element import Attribute, Element

from itsm.base import Base


class ADSMBase(Base):

	ref_sep = ';#'
	batch_size = 20

	def _delayed_adsm_client(self, var, sp):
		if not hasattr(self, var):
			setattr(self, var, self._named_client('SP_ADSM', lambda url, **x: self.create_sharepoint_client(url + '/' + sp, **x)))
		return getattr(self, var)

	# Access web services clients for ADSM endpoints

	@property
	def adsm_lists(self):  return self._delayed_adsm_client('_adsm_lists',  'Lists.asmx?WSDL')
	@property
	def adsm_people(self): return self._delayed_adsm_client('_adsm_people', 'People.asmx?WSDL')
	@property
	def adsm_webs(self):   return self._delayed_adsm_client('_adsm_webs',   'Webs.asmx?WSDL')

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

	def listitem_ref(self, list_uuid, view_uuid, field, field_value, display_field='_ows_Title'):
		cache_key = '%s/%s/%s' % (list_uuid, view_uuid, field)
		table = self.listitem_ref._cache.get(cache_key)
		if not table:
			list_items = self.adsm_lists.service.GetListItems(list_uuid, view_uuid)
			list_items_rows = list_items.listitems.data.row if int(list_items.listitems.data._ItemCount) > 1 \
			          else [list_items.listitems.data.row] if int(list_items.listitems.data._ItemCount) > 0 \
			          else []
			table = dict(map(lambda x: (x[field], x), list_items_rows))
			self.listitem_ref._cache[cache_key] = table
		match = table.get(field_value)
		if match:
			return '%s%s%s' % (match['_ows_ID'], ADSMBase.ref_sep, match[display_field] if display_field else '')
		return None
	listitem_ref._cache = {}

	def listitem_refs(self, list_uuid, view_uuid, field, field_values, display_field='_ows_Title'):
		if not field_values:
			return None
		if isinstance(field_values, basestring):
			field_values = [field_values]
		return ADSMBase.ref_sep.join(filter(lambda x: x, map(lambda y: self.listitem_ref(list_uuid, view_uuid, field, y, display_field=display_field), field_values)))

	def reset_caches(self):
		self.listitem_ref._cache = {}
		self.person_ref._cache = {}

	# Sync functions

	def sync_to_list_by_comparison(self, list_uuid, view_uuid, list_items_compare_key, ext_items, ext_items_compare_key, compare_f, field_map):
		list_items = self.adsm_lists.service.GetListItems(list_uuid, view_uuid)
		list_items_rows = list_items.listitems.data.row if int(list_items.listitems.data._ItemCount) > 1 \
		            else [list_items.listitems.data.row] if int(list_items.listitems.data._ItemCount) > 0 \
		            else []
		list_items_map = dict(map(lambda x: (x[list_items_compare_key], x), list_items_rows))
		
		method_idx = 1
		batch = Element('Batch')\
		       .append(Attribute('OnError', 'Continue'))\
		       .append(Attribute('ListVersion', 1))

		def update(b):
			updates = Element('ns1:updates').append(b)
			return self.adsm_lists.service.UpdateListItems(listName=list_uuid, updates=updates)

		for ext_item in ext_items:
			list_item = list_items_map.get(ext_item[ext_items_compare_key])
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
			          .setText(item_id))
			
			for dst, src in field_map:
				try:
					v = ext_item[src] if isinstance(src, basestring) else src(ext_item)
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

