#!/usr/bin/env python
# coding=utf-8

from suds.sax.element import Attribute, Element

from itsm.adsm import ADSMBase


class DeleteItems(ADSMBase):

	def argument_parser(self):
		parser = super(DeleteItems, self).argument_parser()

		parser.add_argument('list', help='list uuid')
		parser.add_argument('view', help='view uuid')
		parser.add_argument('-d', help='dry run', action='store_true')

		return parser

	def main(self):
		items = self.adsm_lists.service.GetListItems(self.args.list, self.args.view).listitems.data.row

		method_idx = 1
		batch = Element('Batch')\
		       .append(Attribute('OnError', 'Continue'))\
		       .append(Attribute('ListVersion', 1))

		def update(b):
			if not self.args.d:
				updates = Element('ns1:updates').append(b)
				print self.adsm_lists.service.UpdateListItems(listName=self.args.list, updates=updates)

		for item in items:
			method = Element('Method')\
			        .append(Attribute('ID', method_idx))\
			        .append(Attribute('Cmd', 'Delete'))\
			        .append(Element('Field')\
			          .append(Attribute('Name', 'ID'))\
			          .setText(item['_ows_ID']))
			batch.append(method)
			print method
			method_idx += 1

			if len(batch) > 20:
				update(batch)
				batch.detachChildren()

		if len(batch) > 0:
			update(batch)


def main(args=None):
	return DeleteItems(args=args).run()


if __name__ == '__main__':
	main()