#!/usr/bin/env python
# coding=utf-8

import calendar
import re

from bs4 import BeautifulSoup
from urllib import urlopen
from urlparse import urljoin
from pprint import pprint
from string import digits

from itsm.adsm import ADSMBase


class BuildingDetailsToADSM(ADSMBase):

	# Production
	uc_list = '{8E3E8107-9FEF-406F-880E-8C980E7400EE}'
	key_fields = ('ID', 'Title')

	def argument_parser(self):
		parser = super(BuildingDetailsToADSM, self).argument_parser()

		parser.add_argument('buildings_url', default='http://www.ucalgary.ca/facilities/buildings/', nargs='?')
		parser.add_argument('-d', help='dry run', action='store_true')

		return parser

	def main(self):
		months_map = dict((calendar.month_name[i][:3], '%02d' % i) for i in range(1, 13))
		months_map.update(dict((calendar.month_name[i], '%02d' % i) for i in range(1, 13)))

		def conv_date(d):
			if d in ('TBC', '...', '--'):
				return None
			parts = d.split(' ')
			if len(parts) == 1: # Year
				return '%s-01-01 00:00:00' % parts[0]
			elif len(parts) == 2: # Month Year
				return '%s-%s-01 00:00:00' % (parts[1], months_map.get(parts[0], '01'), )
			return None

		field_map = (
			('CIDescription', lambda r: r['Facts'][0] if len(r['Facts']) > 0 else None),
			('CIBusinessOwner', lambda r: self.person_ref(r['Facility Manager'][1])),
			('CIBuildingEmergencyP1', lambda r: filter(lambda x: '(primary)' in x, r['Emergency Assembly Points'])[0] if isinstance(r['Emergency Assembly Points'], list) else r['Emergency Assembly Points']),
			('CIBuildingEmergencyP2', lambda r: filter(lambda x: '(secondary)' in x, r['Emergency Assembly Points'])[0]),
			('CIBuildingArchitect', 'Architect'),
			('CISupplier', 'General Contractor'),
			('CIBuildingZone', 'Zone'),
			('CIBuildingArea', lambda r: r['Building Area'] or r['Original Building Area']),
			('CIDevelopmentDate', lambda r: conv_date(r['Start Date'])),
			('CIReleaseDate', lambda r: conv_date(r['Completion Date'])),
			('CIProductionFundingAmount', 'Cost'),
			('CIResourceURL', lambda r: '%s, %s' % (r['URL'], r['Building'])),
		)

		def rows():
			name_map = {
				'General Services': 'General Services Building',
				'Grounds': 'Grounds Building',
				'Kinesiology Complex': 'Kinesiology A',
				'MacKimmie Library Block and Tower': 'MacKimmie Tower',
				'Materials Handling': 'Materials Handling Facility',
				'Math Sciences Building and Tower': 'Mathematical Sciences',
				'Education Block and Tower': 'Education Tower',
				'Weather Research Station': 'Weather Station',
				'Heritage Medical': 'Heritage Medical Research Building',
				'Priddis Observatory': 'Rothney Astrological Observatory Lab',
				'International House / Hotel Alma': 'International House (Dr. Fok Ying Tong)',
				'Yamnuska': 'Yamnuska Hall'
			}

			for building in self.buildings():
				building['AdjustedBuilding'] = name_map.get(building['Building'], building['Building'])
				yield building

		def compare_f(ext_item, list_item):
			return 'Update' if list_item else None

		self.sync_to_list_by_comparison(BuildingDetailsToADSM.uc_list, None, BuildingDetailsToADSM.key_fields, '_ows_Title', rows(), 'AdjustedBuilding', compare_f, field_map, content_type='Building', fuzzy=True, commit=not self.args.d)

	def buildings(self):
		# Fetch the list of all buildings
		buildings_soup = BeautifulSoup(urlopen(self.args.buildings_url))
		building_anchors = buildings_soup.find_all('a', href=re.compile('/facilities/buildings/'))

		# For each building, fetch its information
		for building_anchor in building_anchors:
			building_url = urljoin(self.args.buildings_url, building_anchor['href'])
			building_soup = BeautifulSoup(urlopen(building_url))

			info_heading = building_soup.find('h2', text='INFORMATION') or building_soup.find_all('h2')[5]
			info_detail_candidates = []
			for sibling in info_heading.next_siblings:
				if sibling.name == 'p':
					info_detail_candidates.append(sibling)
				elif sibling.name == 'h2':
					break
			info_details = max(info_detail_candidates, key=lambda x: len(x))

			info = {}
			key, values = None, []
			def commit_info(key, values):
				if key != None:
					value = values if len(values) > 1 else values[0] if len(values) == 1 else None
					info[key] = value
			for child in info_details.children:
				if child.name == 'strong':
					commit_info(key, values)
					key = child.text.strip(u': \xa0')
					values = []
				elif child.name == None:
					value = child.strip(u' []\xa0')
					if key == u'Start/Completion Date' or key == u'Date':
						dates = value.split('/')
						commit_info(u'Start Date', (dates[0].strip(),))
						commit_info(u'Completion Date', (dates[1].strip(),) if len(dates) > 1 else [])
						key = None
					elif value:
						if key in ('Building Area', 'Zone'):
							value = ''.join(c for c in value if c in digits)
							value = int(value) if value else 0
						elif key in ('Cost'):
							million = 'million' in value
							value = ''.join(c for c in value if c in digits + '.')
							value = float(value) if value else 0.0
							if million:
								value *= 1000000.0
						values.append(value)
				elif key == u'Facility Manager' and child.name == 'a':
					value = child['href']
					if value.startswith('mailto:'):
						value = (values[-1], value[7:])
						values[-1] = value
			commit_info(key, values)

			facts_heading = building_soup.find('h2', text='INTERESTING FACTS')
			facts_details = facts_heading.find_next_sibling('ul')
			facts = [fact.text.strip() for fact in facts_details.find_all('li')]

			commit_info('Facts', facts)

			info['Building'] = building_anchor.text
			info['URL'] = building_url
			yield info


def main(args=None):
	return BuildingDetailsToADSM(args=args).run()


if __name__ == '__main__':
	main()
