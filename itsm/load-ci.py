#!/usr/bin/env python
# coding=utf-8

import logging

from xlrd import open_workbook
from rules.rules import Model
from rules.context import Context

from itsm.adsm import ADSMBase


class LoadCI(ADSMBase):

	# Production
	apps_list = '{994AB157-F9F1-418F-B074-F36C4945395D}'
	apps_view = '{E6CCBE95-E947-4BEA-AD2F-39F795D358A5}'
	# apps_fields = ('AppID', 'Unit', 'ApplicationSystemOwner', 'Categorization', 'Team', '_ModerationComments',
	#                'AdditionalApprovers', 'Aliases', 'AnnualDevelopmentCost', 'AnnualDevelopmentHours', 'AnnualInfrastructureCost',
	#                'LicenseAnnualCost', 'SLARevenue', 'AnnualSupportCost', 'AnnualSupportHours', 'appmonitoring', 'Architecture',
	#                'Attachments', 'BusinessContacts', 'Business_x0020_Owner_x0020_Deleg', 'BusinessOwnerIncumbent',
	#                'BusinessOwnerRole', 'BusinessService', 'BusinessSystemOwnerIncumbent', 'ChangeSchedule', 'Comments',
	#                'ContentType', 'ContractRenewalSchedule', 'Created', 'Author', 'DependenciesDatabase', 'dcfosop', 'Description',
	#                'DependenciesServer', 'devmonitoring', 'Environments', 'UsersNumber', 'FundedBy', 'HostedOffsite', 'Language',
	#                'LicenseRenewalSchedule', 'LifecyclePhase', 'LifecyclePhaseComments', 'Module', 'Origin', 'PortfolioRollup',
	#                'Service_x0020_Release_x0020_Last', 'VendorSignedContract', 'SLASigned', 'DependenciesTechnical',
	#                'TechnicalSupport', 'Type', 'UserCommunities', 'Vendor', '_UIVersionString', 'Developer_x0020_1',
	#                'Devleoper_x0020_2', 'Application_x0020_Authentication', 'Application_x0020_Service_x0020_', 'DR_x0020_Priority',
	#                'Email_x0020_Functions', 'Email_x0020_Integration', 'Implementation_x0020_Date', 'Open_x0020_Ports',
	#                'Product_x0020_Info_x0020_URL', 'Production_x0020_URL', 'PS_x0020_Integration', 'QA_x0020_URL',
	#                'Release_x0020_Status', 'Repository_x0020_URL', 'Shutdown_x0020__x002f__x0020_Sta',
	#                'Support_x0020_Contact_x0020_Info', 'ID', 'Email_x0020_follow_x0020_up_x002', 'Title')

	def argument_parser(self):
		parser = super(LoadCI, self).argument_parser()

		parser.add_argument('phase')
		parser.add_argument('model', default='../../Data/AD_Applications.irl', nargs='?', help='model file')

		return parser

	@property
	def opc(self):
		if not hasattr(self, '_opc'):
			setattr(self, '_opc', self.mssql_client('OPC_DB'))
		return getattr(self, '_opc')

	def prepare_for_main(self):
		self.model = Model.modelFromFile(self.args.model)
		self.ctx = Context(model=self.model)
		self.ctx['phase'] = self.args.phase

	def combine_refs(self, refs):
		return ADSMBase.ref_sep.join(filter(lambda x: x, set(refs)))

	def main(self):
		ctx = self.ctx

		def sw_license(s):
			v = {
				'Commercial (closed)': 'Proprietary License',
				'Commercial (open)': 'Proprietary License',
				'Homegrown': None,
				'Open-source': 'Other Open Source License'
			}.get(s)
			return self.choice_ref('Configuration Item Choice', '_ows_Title', v)

		def unit(s):
			if s.startswith('Faculty of ') or s.startswith('Facutly of '):
				s = s[11:]
			v = {
				'Academic Community': 'Provost and Vice-President (Academic)',
				'Arctic Institute': 'Arctic Institute of North America',
				'Calgary Centre for Clinical Research': None,
				'Career Services & Haskayne Career Services': 'Career Services',
				'Centre for International Students and Study Abroad (CISSA)': 'Centre for International Students and Study Abroad',
				'Deskside Support': 'Information Technologies',
				'EVDS': 'Environmental Design',
				'Education': 'Werklund School of Education',
				'Medicine': 'Cumming School of Medicine',
				'Hotel & Conference Services': 'Hotel Alma',
				'Human Resources (Business Process & PeopleSoft System Training)': 'Human Resources',
				'International Student Center': 'International Student Services (ISS)',
				'Parking Services': 'Parking & Transportation Services',
				'Provost': 'Provost and Vice-President (Academic)',
				'Research': 'Vice-President Research',
				'Risk Management': 'Risk Management and Insurance',
				'Support Centre': 'Information Technologies',
				'Taylor Institute': 'Teaching and Learning Centre'
			}.get(s, s)
			return self.uc_ref('Unit', '_ows_Title', v, fuzzy=True, max_dist=8)

		def lifecycle(s):
			v = {
				'End-of-life': 'Decomission',
				'R&D': 'Plan',
				'Testing/Release': 'Test',
				'Upgrading': 'Sustain',
			}.get(s, s)
			return self.choice_ref('Configuration Item Choice', '_ows_Title', v, fuzzy=True)

		def funding_source(s1, s2):
			if 'Client' in s1 or 'Cleint' in s1 or s1 == 'Cost Recovery':
				return unit(s2)
			else:
				v = {
					'Base (ELT)': 'Office of the President',
					'Base Funded': 'Information Technologies',
					'FM & D': 'Facilities Management',
					'GIS': 'Information Technologies',
					'Internal': 'Information Technologies',
					'IT': 'Information Technologies',
					'Library for Initial Development': 'Libraries and Cultural Resources',
					'Portal Project': 'Information Technologies',
					'Provost Office': 'Provost and Vice-President (Academic)',
					'Unicard Office': 'UNICARD',
					'Vet Med': 'Veterinary Medicine',
					'VP Finance': 'Vice-President Finance & Services',
					'VPFS Les Tochor': 'Finance'
				}.get(s1, s1)
				return self.uc_ref('Unit', '_ows_Title', v, fuzzy=True)

		base_map = (
			('ExternalDateCreated', '_ows_Created'),
			('ExternalDateModified', '_ows_Modified')
		)

		app_base_map = base_map + (
			('CIExternalReference1', '_ows_ID'),
			('CIExternalReference2', '_ows_AppID'),
			('Title', '_ows_Title'),
			('CIShortTitle', '_ows_Aliases'),
			('CIDescription', '_ows_Description'),
			('CIUnit', lambda r: unit(r.get('_ows_Unit'))),
			('CITechnicalOwner', lambda r: self.person_ref(r.get('_ows_ApplicationSystemOwner'), fuzzy=True)),
			('CITechnicalAgents', lambda r: self.combine_refs([r.get('_ows_Developer_x0020_1'), r.get('_ows_Developer_x0020_2')])),
			('CITechnicalChangeApprover', '_ows_AdditionalApprovers'), # IT Change Coordinator
			('CIBusinessOwner', '_ows_BusinessOwnerIncumbent'), # Business Owner
			('CIBusinessAgents', lambda r: self.combine_refs([r.get('_ows_BusinessContacts'), r.get('_ows_BusinessSystemOwnerIncumbent')])), # Business Contact Information, Business Process Owner
			('CIBusinessChangeApprover', '_ows_Business_x0020_Owner_x0020_Deleg'), # Business Change Approver
			('CISupplier', '_ows_Vendor'),
			('CIReleaseDate', '_ows_Implementation_x0020_Date'), # Implementation Date
			('CISoftwareLicense', lambda r: sw_license(r.get('_ows_Origin'))),
			('CIResourceURL', '_ows_Product_x0020_Info_x0020_URL'),
			('CILifecycleState', lambda r: lifecycle(r.get('_ows_LifecyclePhase'))),
		)

		utility_base_map = base_map + (
			('Title', '_ows_BusinessService'),
			('CIShortTitle', '_ows_Aliases'),
			('CIDescription', '_ows_Description'),
			('CIUnit', lambda r: unit(r.get('_ows_Unit'))),
			('CIBusinessOwner', '_ows_BusinessOwnerIncumbent'), # Business Owner
			('CIBusinessAgents', lambda r: self.combine_refs([r.get('_ows_BusinessContacts'), r.get('_ows_BusinessSystemOwnerIncumbent')])), # Business Contact Information, Business Process Owner
			('CIBusinessChangeApprover', '_ows_Business_x0020_Owner_x0020_Deleg'), # Business Change Approver
			('CIReleaseDate', '_ows_Implementation_x0020_Date'), # Implementation Date
			('CILifecycleState', lambda r: lifecycle(r.get('_ows_LifecyclePhase'))),
			('CIProductionFundingSource', lambda r: funding_source(r.get('_ows_FundedBy'), r.get('_ows_Unit'))),
			('CIProductionFundingAmount', '_ows_AnnualDevelopmentCost'),
			('CISustainmentFundingSource', '_ows_FundedBy'),
			('CISustainmentFundingAmount', '_ows_AnnualSupportCost'),
		)

		solution_base_map = base_map + (
			('CIExternalReference1', '_ows_ID'),
			('CIExternalReference2', '_ows_AppID'),
			('Title', '_ows_Title'),
			('CIShortTitle', '_ows_Aliases'),
			('CIDescription', '_ows_Description'),
			('CIUnit', lambda r: unit(r.get('_ows_Unit'))),
			('CITechnicalOwner', lambda r: self.person_ref(r.get('_ows_ApplicationSystemOwner'), fuzzy=True)),
			('CITechnicalAgents', lambda r: self.combine_refs([r.get('_ows_Developer_x0020_1'), r.get('_ows_Developer_x0020_2')])),
			('CITechnicalChangeApprover', '_ows_AdditionalApprovers'), # IT Change Coordinator
			('CIBusinessOwner', '_ows_BusinessOwnerIncumbent'), # Business Owner
			('CIBusinessAgents', lambda r: self.combine_refs([r.get('_ows_BusinessContacts'), r.get('_ows_BusinessSystemOwnerIncumbent')])), # Business Contact Information, Business Process Owner
			('CIBusinessChangeApprover', '_ows_Business_x0020_Owner_x0020_Deleg'), # Business Change Approver
			('CISupplier', '_ows_Vendor'),
			('CIReleaseDate', '_ows_Implementation_x0020_Date'), # Implementation Date
			('CISoftwareLicense', lambda r: sw_license(r.get('_ows_Origin'))),
			('CIResourceURL', '_ows_Product_x0020_Info_x0020_URL'),
			('CILifecycleState', lambda r: lifecycle(r.get('_ows_LifecyclePhase'))),
		)

		phase_maps = {
			'1.1': app_base_map + (

			),

			'1.2': app_base_map + (

			),

			'2.1': utility_base_map + (

			),

			'2.2': utility_base_map + (

			),

			'3.1': solution_base_map + (

			)
		}

		# Get the list of connected applications from OPC.Application
		print 'Loading OPC.Application'
		cursor = self.opc.cursor()
		cursor.execute("SELECT DISTINCT ApplicationName FROM Application WHERE ServerId != 5180")
		opc_apps = [r[0] for r in cursor]
		ctx['opc_apps'] = opc_apps
		print '   Done OPC.Application'

		# Get the list items from AD Applications
		print 'Loading AD Applications'
		list_items = self.adsm_lists.service.GetListItems(LoadCI.apps_list, LoadCI.apps_view, rowLimit="2").listitems.data.row
		# list_items = self.adsm_lists.service.GetListItems(LoadCI.apps_list, LoadCI.apps_view).listitems.data.row
		print '   Done AD Applications'
		def rows():
			for row in list_items:
				ctx['row'] = row
				if ctx['qualifyingRow']:
					# print str(row.__dict__)
					yield row.__dict__

		if ctx['phase'].startswith('2.'):
			scanned_titles = set()
		def compare_f(ext_item, list_item):
			print list_item

			if ctx['phase'].startswith('2.') and ext_item['_ows_BusinessService'] in scanned_titles:
				return None
			if list_item == None:
				if ctx['phase'].startswith('2.'):
					scanned_titles.add(ext_item['_ows_BusinessService'])
				return 'New'
			else:
				pass

			return None

		# logging.basicConfig(level=logging.INFO)
		# logging.getLogger('suds.client').setLevel(logging.DEBUG)

		self.sync_to_list_by_comparison(self.ci_list_uuid, None, ('ID', 'Title', 'CIExternalReference1'), '_ows_CIExternalReference1', rows(), '_ows_ID', compare_f, phase_maps[ctx['phase']], content_type=unicode(ctx['contentType'], 'utf-8'), folder=ctx['folder'], commit=not self.args.d)


def main(args=None):
	return LoadCI(args=args).run()


if __name__ == '__main__':
	main()
