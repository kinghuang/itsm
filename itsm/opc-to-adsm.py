#!/usr/bin/env python
# coding=utf-8

import datetime
import logging
import re
import time

from suds.sax.parser import Parser

from itsm.adsm import ADSMBase


class OPCToADSM(ADSMBase):

	@property
	def opc(self):
		if not hasattr(self, '_opc'):
			setattr(self, '_opc', self.mssql_client('OPC_DB'))
		return getattr(self, '_opc')

	def prepare_for_main(self):
		# logging.basicConfig(level=logging.INFO)
		# logging.getLogger('suds.client').setLevel(logging.DEBUG)
		pass

	def main(self):
		# Connect to the OPC database and fetch columns for the Server table
		cursor = self.opc.cursor()
		cursor.execute("SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Server'")
		server_columns = [row[3] for row in cursor]

		def extract_person_ref(s):
			result = re.search(r'[-0-9a-zA-Z.+_]+@[-0-9a-zA-Z.+_]+\.[a-zA-Z]{2,4}', s)
			return self.person_ref(result.group(0)) if result else None

		def extract_people_refs(ss):
			return ADSMBase.ref_sep.join(filter(lambda x: x, map(extract_person_ref, ss)))

		def building(s):
			if s == 'Calg Ctr Innovative Tech':
				s = 'Calgary Centre for Innovative Technology'
			return self.uc_ref('Building', '_ows_Title', s, fuzzy=True, max_dist=10)

		def site(s):
			s = {
				'South Campus': 'Foothills Campus',
				'Spy Hill': 'Spyhill Campus',
				'UofA': 'University of Alberta'
			}.get(s, s)
			return self.uc_ref('Site', '_ows_Title', s, fuzzy=True)

		def status(s):
			s = {
				'Development': 'Develop',
				'Pre-Production': 'Deploy',
				'Production': 'Sustain',
				'Production Standby': 'Sustain',
				'Testing': 'Test'
			}.get(s, s)
			return self.choice_ref('Configuration Item Choice', '_ows_Title', s, fuzzy=True)

		# Field maps
		base_map = (
			('Title', 'ServerName'),
			('CIDescription', lambda r: ' '.join(filter(lambda x: x, (r['Function1'], r['Function2'])))),
			('CITechnicalOwner', lambda r: extract_person_ref(r['OwnerContact'])),
			('CITechnicalAgents', lambda r: extract_people_refs(map(lambda x: r[x], ('SystemAdmin1', 'SystemAdmin2', 'SystemAdmin3')))),
			('CIBusinessOwner', lambda r: extract_person_ref(r['BusinessOwner'])),
			('CIReleaseDate', lambda r: r['InstallDate'].isoformat().replace('T', ' ')),
			('CILifecycleState', lambda r: status(r['Status'])),
			('CIHostName', 'HostName'),
			('CISupplier', 'Make'),
			('CIModel', 'Model'),
			('CISite', lambda r: site(r['SiteGroup'])),
			('CIBuilding', lambda r: building(r['Site'])),
			('CIServerZone', lambda r: '-'.join(r['ZoneLocation'].split('-', 2)[:2])),
			('ExternalDateModified', lambda r: r['ModifiedDate'].isoformat().replace('T', ' ')),
			('CIExternalReference1', 'ServerId'),
			('CIExternalReference2', 'UCTagNum')
		)

		physical_servers_map = base_map + (
			# ('CITypeAsset', 'Tier3'),
		)

		virtual_machines_map = base_map + (
			# ('CITypeAsset', 'Tier3'),
		)

		# Fetch relevant servers from OPC and create yield function
		# zipping together results with column names
		def physical_servers():
			cursor.execute("SELECT * FROM Server WHERE Tier2 = 'Computing Device' AND Tier3 = 'Server' AND Status NOT IN ('Removed', 'Research')")
			# row = cursor.fetchone()
			# row = (str(row[0]),) + row[1:]
			# yield dict(zip(server_columns, row))
			for row in cursor:
				yield dict(zip(server_columns, (str(row[0]),) + row[1:]))

		def virtual_machines():
			cursor.execute("SELECT * FROM Server WHERE Tier2 = 'Computing Device' AND Tier3 = 'Virtual Host' AND Status NOT IN ('Removed', 'Research')")
			# row = cursor.fetchone()
			# row = (str(row[0]),) + row[1:]
			# yield dict(zip(server_columns, row))
			for row in cursor:
				yield dict(zip(server_columns, (str(row[0]),) + row[1:]))

		# Comparison function for creating and updating servers
		def compare_f(ext_item, list_item):
			if list_item == None:
				return 'New'
			else:
				list_item_date = datetime.datetime(*time.strptime(list_item['_ows_ExternalDateModified'], '%Y-%m-%d %H:%M:%S')[:6])
				if ext_item['ModifiedDate'] > list_item_date:
					return 'Update'
			return None

		fields = ('ID', 'Title', '_ows_CIExternalReference1', '_ows_ExternalDateModified')
		folder = '/sites/ADSM/Lists/CI/Assets'

		# Sync
		self.sync_to_list_by_comparison(self.ci_list_uuid, None, fields, '_ows_CIExternalReference1', physical_servers(), 'ServerId', compare_f, physical_servers_map, content_type=u'Asset—Server', folder=folder, commit=not self.args.d)
		self.sync_to_list_by_comparison(self.ci_list_uuid, None, fields, '_ows_CIExternalReference1', virtual_machines(), 'ServerId', compare_f, virtual_machines_map, content_type=u'Alias—Virtual Machine', folder=folder, commit=not self.args.d)


def main(args=None):
	return OPCToADSM(args=args).run()


if __name__ == '__main__':
	main()