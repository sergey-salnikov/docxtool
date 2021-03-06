import argparse
import re

import docx

docx.oxml.ns.nsmap['v'] = 'urn:schemas-microsoft-com:vml'


def main():
    args = parse_args()
    for file in args.file:
        document = docx.Document(file)
        document.element.xpath
        if args.resize:
            what, how = args.resize.split(':')
            elements = find_elements(document, what)
            resizer = parse_resize_spec(how)
            for element in elements:
                resize_element(element, resizer)
            print("Resized %d elements" % len(elements))
        if not args.test:
            document.save(file)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--resize', metavar='SPEC')
    parser.add_argument('-t', '--test', action='store_true')
    parser.add_argument('file', nargs='+')
    return parser.parse_args()


def find_elements(document, query):
    if query in ['formulae', 'formulas']:
        # I'm sure this would break on other documents. Worked on the
        # one that motivated this script.
        return document.element.xpath('//w:object/v:shape')
    raise ValueError("Couldn't parse query: '%s'" % query)


def parse_resize_spec(spec):
    if spec[-1] == '%':
        multiplier = float(spec[:-1]) / 100
        return lambda width, height: (multiplier * width, multiplier * height)
    raise ValueError("Couldn't parse resize spec: '%s'" % spec)


def resize_element(element, resizer):
    style = element.attrib['style']
    old_width = pt_in_style(style, 'width')
    old_height = pt_in_style(style, 'height')
    new_width, new_height = resizer(old_width, old_height)
    style = pt_in_style(style, 'width', new_width)
    style = pt_in_style(style, 'height', new_height)
    element.attrib['style'] = style

    for position_xml in element.getparent().getparent().xpath('w:rPr/w:position'):
        attrib_name = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'
        position = int(position_xml.attrib[attrib_name])

        # This is measured in OOXML in strange units: 1/2 pt. So, this
        # really adjusts by half of height difference, which should
        # keep the center of the formula at the same height. One would
        # probably like to keep the baseline instead, but it doesn't
        # seem to be recorded anywhere: when I resize the formula in
        # Word manually, it keeps the bottom line of the formula, just
        # like the version of this script without the hack.
        position -= round(new_height - old_height)
        position_xml.attrib[attrib_name] = str(position)


def pt_in_style(style, attribute, new_value=None):
    pattern = '(.*%s:)([0-9.]+)(pt(?:|;.*))$' % (attribute,)
    match = re.match(pattern, style)
    if new_value:
        replacement = str(round(new_value, 3))
        return match.group(1) + replacement + match.group(3)
    else:
        return float(match.group(2))
