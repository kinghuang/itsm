#!/usr/bin/env python
# coding=utf-8

from itsm.adsm import ADSMBase


class UnitsToADSM(ADSMBase):

	# Production
	units_list = '{3BEA7DBB-27C7-45DF-A187-A589CA062600}'
	key_fields = ('ID', 'Title', 'CIExternalReference1', 'CIExternalReference2')

	# Development
	# units_list = '{C82DE52C-9AA2-43B7-B739-5047A002A15A}'
	
	@property
	def unitis_units(self):
		if not hasattr(self, '_unitis_units'):
			setattr(self, '_unitis_units', self.webservices_client('UNITIS_UNITS'))
		return getattr(self, '_unitis_units')
	
	def main(self):
		# Field maps
		create_update_field_map = (
			('Title', 'Name'),
			('CIExternalReference1', 'Id'),
			('CIExternalReference2', 'EntityVersion'),
			('UNITISUnitCode', 'Code'),
			('UNITISUnitType', 'Type'),
			('UNITISUnitWebsite', 'Website'),
			('UNITISUnitKeywords', 'Keywords'),
			('UNITISUnitBuilding', lambda r: r['Rooms']['ContactRoom'][0]['Room']['BuildingName']),
			('UNITISUnitBuildingCode', lambda r: r['Rooms']['ContactRoom'][0]['Room']['BuildingCode'].upper()),
			('UNITISUnitRoomNumber', lambda r: r['Rooms']['ContactRoom'][0]['Room']['Number']),
			('UNITISUnitPhone', lambda r: '+%s (%s) %s' % (r['Phones']['ContactPhone'][0]['Phone']['CountryCode'], r['Phones']['ContactPhone'][0]['Phone']['AreaCode'], r['Phones']['ContactPhone'][0]['Phone']['Number'])),
			('UNITISUnitEmail', lambda r: r['Emails']['ContactEmail'][0]['Address']),
			('UNITISUnitCoordinators', lambda r: ';#'.join(map(lambda x: self.person_ref(x['Email']), r['Coordinators']['Coordinator'])))
		)

		parents_children_field_map = (
			('UNITISUnitParentUnits', lambda r: self.listitem_refs(UnitsToADSM.units_list, None, UnitsToADSM.key_fields, '_ows_CIExternalReference1', [x['Id'] for x in r['Parents']['Unit']])),
			('UNITISUnitChildUnits', lambda r: self.listitem_refs(UnitsToADSM.units_list, None, UnitsToADSM.key_fields, '_ows_CIExternalReference1', [x['Id'] for x in r['Children']['Unit']]))
		)

		# Fetch all public units from UNITIS
		public_units = self.unitis_units.service.GetPublicUnits(getAliases=False).PublicUnit
		updated_unit_ids = set()

		# Comparison function for creating and updating units
		def create_update_compare_f(ext_item, list_item):
			if list_item == None:
				updated_unit_ids.add(ext_item['Id'])
				return 'New'
			elif int(ext_item['EntityVersion']) > int(list_item['_ows_CIExternalReference2']):
				updated_unit_ids.add(ext_item['Id'])
				return 'Update'
			return None

		# Comparison function for linking units to their parents and children
		def parents_children_compare_f(ext_item, list_item):
			return 'Update' if ext_item['Id'] in updated_unit_ids else None

		# Sync
		self.sync_to_list_by_comparison(UnitsToADSM.units_list, None, UnitsToADSM.key_fields, '_ows_CIExternalReference1', public_units, 'Id', create_update_compare_f, create_update_field_map)
		self.reset_caches()
		self.sync_to_list_by_comparison(UnitsToADSM.units_list, None, UnitsToADSM.key_fields, '_ows_CIExternalReference1', public_units, 'Id', parents_children_compare_f, parents_children_field_map)


def main(args=None):
	return UnitsToADSM(args=args).run()


if __name__ == '__main__':
	main()
