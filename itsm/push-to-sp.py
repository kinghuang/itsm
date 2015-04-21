#!/usr/bin/env python
# coding=utf-8

import json
import logging
import sys

from suds.sax.element import Attribute, Element
from xlrd import open_workbook

from itsm.base import Base


class PushToSP(Base):

	def argument_parser(self):
		parser = super(PushToSP, self).argument_parser()

		parser.add_argument('op', help='operation to perform (new, update, delete)')
		parser.add_argument('type', help='type of thing to process (columns, content-type, list)')
		parser.add_argument('spec', help='argument to the type (identifier or group)')
		parser.add_argument('spreadsheet', help='path to Excel spreadsheet containing specifications', default='../Spreadsheets/ITSM Field Definitions.xlsx', nargs='?')
		parser.add_argument('sheet', help='sheet name and optional sub-type in spreadsheet')

		return parser

	def prepare_for_main(self):
		# logging.basicConfig(level=logging.INFO)
		# logging.getLogger('suds.client').setLevel(logging.DEBUG)

		self._sp_webs = None
		self._sp_lists = None

	@property
	def sp_webs(self):
		if not self._sp_webs:
			self._sp_webs = self.sharepoint_client('SP_ADSM_WEBS')
		return self._sp_webs

	@property
	def sp_lists(self):
		if not self._sp_lists:
			self._sp_lists = self.sharepoint_client('SP_ADSM_LISTS')
		return self._sp_lists

	def main(self):
		# Get the Excel workbook and sheet.
		# The sheet argument might have a slash with a sub-type name. Split as needed.
		sheet_name, sub_type = self.args.sheet.split('/', 1) if '/' in self.args.sheet else (self.args.sheet, None)
		workbook = open_workbook(self.args.spreadsheet)
		sheet = workbook.sheet_by_name(sheet_name)

		# Build mapping of spreadsheet column names to index
		header_row = sheet.row(0)
		cols = dict([(column.value, index) for index, column in zip(range(0, len(header_row)), header_row)])

		# If we are processing content-type or columns, fetch the site columns
		if self.args.type in ('content-type', 'columns'):
			sp_columns = {f._Name:f for f in self.sp_webs.service.GetColumns().Fields.Field}
			sp_source = sp_columns

		# If we are processing a content-type, fetch the existing type
		if self.args.type == 'content-type':
			sp_ctype = self.sp_webs.service.GetContentType(self.args.spec)
			sp_ctype_fieldnames = set(f._Name for f in sp_ctype.ContentType.Fields.Field)
			sp_source = sp_ctype_fieldnames

		# If we are processing a list, fetch the existing list
		if self.args.type == 'list':
			sp_list = self.sp_lists.service.GetList(self.args.spec)
			sp_list_fieldnames = set(f._Name for f in sp_list.List.Fields.Field)
			sp_source = sp_list_fieldnames

		# Holds fields for the current operation
		op_fields = Element('Fields')

		for row_idx in range(1, sheet.nrows):
			row = sheet.row(row_idx)
			c = lambda x: row[cols[x]].value

			# Skip rows that do not match sub-type, or rows with a sub-type if one is
			# not defined.
			if (sub_type and c('Content Type') != sub_type) or (not sub_type and c('Content Type')):
				continue

			# Skip rows that do not have an internal name. An internal name is required
			# to reference in SharePoint.
			if not c('Field Internal Name'):
				continue

			# If creating, check that the row doesn't already exist
			if self.args.op == 'new' and c('Field Internal Name') in sp_source:
				continue

			sp_args = json.loads(c('Type Arguments')) if c('Type Arguments')else {}
			if sp_args.get('SPDefined', False):
				continue

			method = Element('Method').append(Attribute('ID', row_idx))
			field = Element('Field')

			if self.args.op == 'delete':
				# if sp_args.get('SPDelete', False) and c('Field Internal Name') in sp_source:
				if c('Field Internal Name') in sp_source:
					if self.args.type == 'columns' and sp_source[c('Field Internal Name')]._Group != self.args.spec:
						continue
					field.set('Name', c('Field Internal Name'))
				else:
					continue
			else:
				attrs = [
					('DisplayName', c('Field Internal Name') if self.args.op is 'new' else c('Field Display Name')),
					('Name', c('Field Internal Name')),
					('Description', c('Description')),
					('Type', c('Type')),
					('List', c('Values') if c('Type') == 'Lookup' else None)
				]

				if not c('Type'):
					sys.exit('Missing type for ' + c('Field Display Name'))

				if self.args.type == 'columns':
					attrs.append(('Group', self.args.spec))
				elif self.args.type == 'content-type':
					site_col = sp_columns.get(c('Field Internal Name'))
					attrs.append(('ID', site_col['_ID']))
					attrs.append(('SourceID', site_col['_SourceID']))

				for att, value in attrs:
					if value:
						field.set(att, value)

				if c('Type') in ('Choice', 'MultiChoice'):
					choices = Element('CHOICES')
					for c in c('Values').splitlines(False):
						choices.append(Element('CHOICE').setText(c))
					field.append(choices)

				if sp_args:
					for k, v in sp_args.items():
						if k in ('Default', ):
							field.append(Element(k),setText(v))
						else:
							field.set(k, v)

			method.append(field)
			op_fields.append(method)

		print '---Fields---'
		print op_fields

		if not op_fields.isempty():
			v = self.args.op + 'Fields'

			if self.args.type == 'columns':
				print self.sp_webs.service.UpdateColumns(**{v:Element('ns1:%s' % v).append(op_fields)})
			elif self.args.type == 'content-type':
				print self.sp_webs.service.UpdateContentType(contentTypeId=self.args.spec, **{v:Element('ns1:%s' % v).append(op_fields)})
			elif self.args.type == 'list':
				print self.sp_lists.service.UpdateList(listName=self.args.spec, **{v:Element('ns1:%s' % v).append(op_fields)})


def main(args=None):
	return PushToSP(args=args).run()


if __name__ == '__main__':
	main()
