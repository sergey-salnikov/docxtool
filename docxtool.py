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
        document.save(file)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('--resize', metavar='SPEC')
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
    width = pt_in_style(style, 'width')
    height = pt_in_style(style, 'height')
    width, height = resizer(width, height)
    style = pt_in_style(style, 'width', width)
    style = pt_in_style(style, 'height', height)
    element.attrib['style'] = style


def pt_in_style(style, attribute, new_value=None):
    pattern = '(.*%s:)([0-9.]+)(pt(?:|;.*))$' % (attribute,)
    match = re.match(pattern, style)
    if new_value:
        replacement = str(round(new_value, 3))
        return match.group(1) + replacement + match.group(3)
    else:
        return float(match.group(2))
