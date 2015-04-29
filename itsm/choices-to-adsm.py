#!/usr/bin/env python
# coding=utf-8

from xlrd import open_workbook

from itsm.adsm import ADSMBase


class Choices(ADSMBase):

	choices_list = '{AD9E7317-ABEC-4250-B538-78EC12829734}'
	ref_view = '{86478DC7-EF04-421C-A63D-97944764D897}'

	def argument_parser(self):
		parser = super(Choices, self).argument_parser()

		parser.add_argument('spreadsheet', help='Excel spreadsheet', default='../../Spreadsheets/Choices.xlsx', nargs='?')
		parser.add_argument('sheet', help='sheet name')
		parser.add_argument('-d', help='dry run', action='store_true')

		return parser

	def main(self):
		workbook = open_workbook(self.args.spreadsheet)
		sheet = workbook.sheet_by_name(self.args.sheet)

		header_row = sheet.row(0)
		columns = [col.value for col in header_row]

		field_map = (
			('Title', 'Title'),
			('CIShortTitle', 'Short Title'),
			('CIResourceURL', lambda r: '%s, %s' % (r['Resource URL'], r['Resource Title']) if r['Resource URL'] else None),
			('CIDescription', 'Description'),
			('CIChoiceType', lambda r: self.args.sheet)
		)

		def compare_f(ext_item, list_item):
			if list_item == None:
				return 'New'
			elif list_item['_ows_CIChoiceType'] == self.args.sheet:
				return 'Update'
			return None

		def rows():
			for row_idx in range(1, sheet.nrows):
				row = sheet.row(row_idx)
				yield dict(zip(columns, (x.value for x in row)))

		self.sync_to_list_by_comparison(Choices.choices_list, None, ('ID', 'Title'), '_ows_Title', rows(), 'Title', compare_f, field_map, content_type='Configuration Item Choice', commit=not self.args.d)


def main(args=None):
	return Choices(args=args).run()


if __name__ == '__main__':
	main()
