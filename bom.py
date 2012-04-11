import xlrd
import json
import urllib2

class BOMParseError (Exception):
    """Raised when there is a problem parsing the BOM."""
    def __init__(self, row):
        self.line = row
        self.m = "There was an error parsing the BOM on row %d."%(row)

class BOMRequestError (Exception):
    """Raised when there is a problem with the BOM request."""
    def __init__(self, httpError, url):
        self.httpError = httpError
        self.url = url
        self.m = "There was an error requesting BOM information from Octopart."

class OctopartBOMInfo:
    """This class represents the data from Octopart. Each line from the original
    BOM is a key to a list of part data from Octopart."""
    def __init__(self, bom):
        self.bom = bom
        self.octopart_data = None
    def construct_line(self, line):
        """Convert a BOMLine into a dictionary."""
        dict_line = {}
        if len(line.manufacturer) > 0:
            dict_line["manufacturer"] = line.manufacturer
        if len(line.mpn) > 0:
            dict_line["mpn"] = line.mpn
        if len(line.supplier) > 0:
            dict_line["supplier"] = line.supplier
        if len(line.sku) > 0:
            dict_line["sku"] = line.sku

        return dict_line

    def construct_lines_list(self):
        """Convert a BOM into a list of lines."""
        return [self.construct_line(line) for line in self.bom.lines]
            
    def retrieve_octopart_data(self):
        """Send the BOM data to octopart and store the retrieved data in 
        self.octopart_data"""
        json_lines = json.dumps(self.construct_lines_list(), separators=(',', ':'))
        url = "http://octopart.com/api/v2/bom/match?lines=%s"%(json_lines)
        try:
            self.octopart_data = urllib2.urlopen(url).read()
        except urllib2.HTTPError, e:
            raise BOMRequestError(e, url)

class BOM:
    """This class represents a bill of materials."""
    def __init__(self):
        self.lines = []
    def add_line(self, bom_line):
        """Add a BOMLine."""
        self.lines.append(bom_line)

class BOMLine:
    """This class represents a line from a BOM, including the manufacturer,
    MPN, supplier, and SKU."""
    def __init__(self):
        self.manufacturer = ""
        self.mpn = ""
        self.supplier = ""
        self.sku = ""
    def is_empty(self):
        """Returns True if all of the attributes are empty or null."""
        return not bool(self.manufacturer or self.mpn or self.supplier or self.sku)

class ExcelBOMParser:
    """This class represents a BOM parser. It parses an Excel file and extracts
    lines from it."""
    def __init__(self, *args, **kwargs):
        self.firstline = kwargs['firstline']
        self.manufacturer_col = kwargs['manufacturer_col']
        self.mpn_col = kwargs['mpn_col']
        self.supplier_col = kwargs['supplier_col']
        self.sku_col = kwargs['sku_col']

    def parse_file(self, excel_file_name, sheet_name):
        """Parse the specified sheet of the specified Excel file using the
        options that were given when the parser was created. Returns a list
        of BOMLine objects."""
        book = xlrd.open_workbook(excel_file_name)
        sheet = book.sheet_by_name(sheet_name)

        bom = BOM()

        for row_number in xrange(self.firstline - 1, sheet.nrows):
            try:
                row = sheet.row(row_number)
                bom_line = BOMLine()
                if self.manufacturer_col:
                    bom_line.manufacturer = row[self.manufacturer_col].value
                if self.mpn_col:
                    bom_line.mpn = row[self.mpn_col].value
                if self.supplier_col:
                    bom_line.supplier = row[self.supplier_col].value
                if self.sku_col:
                    bom_line.sku = row[self.sku_col].value
                if not bom_line.is_empty() and (bom_line.supplier.lower() == "digikey" or bom_line.supplier.lower() == "mouser"):
                    bom.add_line(bom_line)
            except Exception, e:
                raise BOMParseError(row_number)

        return bom
        
