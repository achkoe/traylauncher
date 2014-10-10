# -*- coding: utf-8  -*-
# Copyright 2013, Achim Köhler
# All rights reserved, see accompanied file license.txt for details.

# $REV$

import argparse
import traylauncher

if __name__ == "__main__":
	args = argparse.Namespace()
	args.notray = False
	traylauncher.start(args)