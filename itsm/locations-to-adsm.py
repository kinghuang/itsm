#!/usr/bin/env python
# coding=utf-8

import logging

from xlrd import open_workbook

from itsm.adsm import ADSMBase


class LocationsToADSM(ADSMBase):

	# Production
	uc_list = '{8E3E8107-9FEF-406F-880E-8C980E7400EE}'
	key_fields = ('ID', 'Title', 'CISiteCode', 'CIExternalReference1')
		
	def argument_parser(self):
		parser = super(LocationsToADSM, self).argument_parser()

		parser.add_argument('spreadsheet', default='../../Data/Buildings.xlsx', nargs='?')
		parser.add_argument('buildings_sheet', default='Buildings', nargs='?')
		parser.add_argument('sites_sheet', default='Sites', nargs='?')
		parser.add_argument('-d', help='dry run', action='store_true')

		return parser

	def main(self):
		# logging.basicConfig(level=logging.INFO)
		# logging.getLogger('suds.client').setLevel(logging.DEBUG)

		# Field maps
		building_field_map = (
			('Title', 'Building Name'),
			('CIShortTitle', 'Abbreviation'),
			('CISite', lambda r: self.listitem_ref(LocationsToADSM.uc_list, None, ('CIShortTitle', 'ID', 'Title'), '_ows_CIShortTitle', r['Site Code'])),
			('CIExternalReference1', 'Number'),
			('WorkAddress', 'Municipal Address'),
			('WorkCity', lambda r: 'Priddis' if 'Priddis' in r['Municipal Address'] else 'Calgary'),
			('WorkState', lambda r: 'Alberta'),
			('WorkCountry', lambda r: 'Canada')
		)

		site_field_map = (
			('Title', 'Site Name'),
			('CIShortTitle', 'Site Code'),
			('CIResourceURL', 'Map URL')
		)

		workbook = open_workbook(self.args.spreadsheet)
		buildings_sheet = workbook.sheet_by_name(self.args.buildings_sheet)
		sites_sheet = workbook.sheet_by_name(self.args.sites_sheet)

		building_headers = [c.value for c in buildings_sheet.row(0)]
		site_headers = [c.value for c in sites_sheet.row(0)]

		def rows(headers, sheet):
			for row_idx in range(1, sheet.nrows):
				row = sheet.row(row_idx)
				yield dict(zip(headers, (c.value for c in row)))

		def compare_f(ext_item, list_item):
			if list_item == None:
				return 'New'
			else:
				return 'Update'

		# Sync
		self.sync_to_list_by_comparison(LocationsToADSM.uc_list, None, LocationsToADSM.key_fields, '_ows_CIShortTitle', rows(site_headers, sites_sheet), 'Site Code', compare_f, site_field_map, content_type='Site', commit=not self.args.d)
		self.sync_to_list_by_comparison(LocationsToADSM.uc_list, None, LocationsToADSM.key_fields, '_ows_CIExternalReference1', rows(building_headers, buildings_sheet), 'Number', compare_f, building_field_map, content_type='Building', commit=not self.args.d)


def main(args=None):
	return LocationsToADSM(args=args).run()


if __name__ == '__main__':
	main()
