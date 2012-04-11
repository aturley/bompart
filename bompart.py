import sys
from optparse import OptionParser

from bom import BOMParseError, BOMRequestError, OctopartBOMInfo, BOM, ExcelBOMParser

"""This script reads in a Excel file that represents a BOM, parses out the 
relevant rows, sends the data to octopart, and prints the resulting parts
list to stdout."""

def get_option_parser():
    """Return a parser that parses the following command-line options:
    --manufacturer=MANUFACTURER_COLUMN
    --mfg=MANUFACTURER_COLUMN
    --mbn=MBN_COLUMN
    --supplier=SUPPLIER_COLUMN
    --sku=SKU_COLUMN
    --firstline=FIRST_LINE
    """
    usage = "usage: %prog EXCEL_FILE BOM_SHEET [options]"
    parser = OptionParser(usage=usage)
    parser.add_option("--manufacturer", "--mfg", dest="manufacturer",
                      type="int", help="manufacturer column",
                      metavar="MANUFACTURER_COLUMN")

    parser.add_option("--mpn", dest="mpn", type= "int",
                      help="mpn column", metavar="MPN_COLUMN", )

    parser.add_option("--supplier", dest="supplier", type="int",
                      help="supplier column", metavar="SUPPLIER_COLUMN")

    parser.add_option("--sku", dest="sku", type="int",
                      help="sku column", metavar="SKU_COLUMN")

    parser.add_option("--firstline", dest="firstline", type="int",
                      default=0, help="first part line", metavar="FIRST_LINE")

    return parser

def main(argv):
    parser = get_option_parser()

    (options, args) = parser.parse_args(argv)

    if (len(args) < 3):
        parser.print_help()
        sys.exit(1)

    bom_file = args[1]
    bom_sheet = args[2]

    bom_parser = ExcelBOMParser(firstline = options.firstline,
                                manufacturer_col=options.manufacturer,
                                mpn_col=options.mpn,
                                supplier_col=options.supplier,
                                sku_col=options.sku)

    try:
        bom = bom_parser.parse_file(bom_file, bom_sheet)

        obi = OctopartBOMInfo(bom)

        obi.retrieve_octopart_data()
        print obi.octopart_data

    except BOMParseError, e:
        print "Error: %s"%(e.m)
        sys.exit(1)

    except BOMRequestError, e:
        print "Error: %s"%(e.m)
        sys.exit(1)

if __name__ == "__main__":
    main(sys.argv)
