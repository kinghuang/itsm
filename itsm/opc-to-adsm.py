#!/usr/bin/env python
# coding=utf-8

import datetime
import re
import time

from itsm.adsm import ADSMBase


class OPCToADSM(ADSMBase):

	# Production
	ci_list = '{83033962-8552-4113-81BC-E31D656D8C3E}'
	key_view = '{B273210C-0218-47F9-AC4B-4EEEE7916E50}'
	folder = '/sites/ADSM/Lists/CI/Assets'

	@property
	def opc(self):
		if not hasattr(self, '_opc'):
			setattr(self, '_opc', self.mssql_client('OPC_DB'))
		return getattr(self, '_opc')

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

		# Field maps
		server_map = (
			('Title', 'ServerName'),
			('CIDescription', lambda r: ' '.join(filter(lambda x: x, (r['Function1'], r['Function2'])))),
			('CITechnicalOwner', lambda r: extract_person_ref(r['OwnerContact'])),
			('CITechnicalAgents', lambda r: extract_people_refs(map(lambda x: r[x], ('SystemAdmin1', 'SystemAdmin2', 'SystemAdmin3')))),
			('CIBusinessOwner', lambda r: extract_person_ref(r['BusinessOwner'])),
			('CIReleaseDate', lambda r: r['InstallDate'].isoformat().replace('T', ' ')),
			('CILifecycleState', lambda r: None),	# Status
			('OPCServerHostName', 'HostName'),
			('OPCServerDeviceType', 'Tier3'),
			('CISupplier', 'Make'),
			('CIModel', 'Model'),
			('OPCServerSiteGroup', 'SiteGroup'),
			('OPCServerSite', 'Site'),
			('OPCServerZone', lambda r: '-'.join(r['ZoneLocation'].split('-', 2)[:2])),
			('OPCServerStatus', 'Status'),
			('CIExternalDateModified', lambda r: r['ModifiedDate'].isoformat().replace('T', ' ')),
			('CIExternalReference1', 'ServerId')
		)

		# Fetch relevant servers from OPC and create yield function
		# zipping together results with column names
		cursor.execute("SELECT * FROM Server WHERE Tier2 = 'Computing Device' AND Tier3 IN ('Server', 'Virtual Host') AND Status NOT IN ('Removed', 'Research')")
		def servers():
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
				list_item_date = datetime.datetime(*time.strptime(list_item['_ows_CIExternalDateModified'], '%Y-%m-%d %H:%M:%S')[:6])
				if ext_item['ModifiedDate'] > list_item_date:
					return 'Update'
			return None

		# Sync
		self.sync_to_list_by_comparison(OPCToADSM.ci_list, OPCToADSM.key_view, '_ows_CIExternalReference1', servers(), 'ServerId', compare_f, server_map, content_type=u'Asset—OPC Server', folder=OPCToADSM.folder)


def main(args=None):
	return OPCToADSM(args=args).run()


if __name__ == '__main__':
	main()