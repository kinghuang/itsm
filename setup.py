#!/usr/bin/env python
# coding=utf-8

import os
import sys

from setuptools import setup, find_packages


if not hasattr(sys, 'version_info') or sys.version_info < (2, 7, 0, 'final'):
	sys.exit('itsm requires Python 2.7 or later.')

setup(
	name = 'itsm',
	version = '0.1',

	author = 'King Chung Huang',
	author_email = 'kchuang@ucalgary.ca',
	description = 'Utilities for ADSM site.',

	packages = find_packages(),
	include_package_data = True,
	install_requires = [

	],
	
	# entry_points = {
	# 	'console_scripts': [
	# 		'push-to-sp = itsm.push-to-sp:main'
	# 	]
	# },

	zip_safe = True
)
