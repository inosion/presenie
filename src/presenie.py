import argparse
import sys
from pptx_tools import PPTXTools
from pptx_merger import PPTXMerger

VERSION = "1.0.0"

def main():
    parser = argparse.ArgumentParser(description=f"Presenie {VERSION} (c) 2024 Inosion")
    parser.add_argument('--verbose', action='store_true', help='Enable verbose output')

    subparsers = parser.add_subparsers(dest='command')

    # List subcommand
    list_parser = subparsers.add_parser('list', help='List slide layouts')
    list_parser.add_argument('--template', type=str, required=True, help='The template to merge data with')

    # Merge subcommand
    merge_parser = subparsers.add_parser('merge', help='Merge data with template')
    merge_parser.add_argument('--template', type=str, required=True, help='The template to merge data with')
    merge_parser.add_argument('--data', type=str, required=True, help='The data file to merge with the template')
    merge_parser.add_argument('--config', type=str, help='Config file to drive merging')
    merge_parser.add_argument('--outFile', type=str, required=True, help='Output filename')

    # Clone subcommand
    clone_parser = subparsers.add_parser('clone', help='Clone PPT slides')
    clone_parser.add_argument('--template', type=str, required=True, help='The template to merge data with')
    clone_parser.add_argument('--outFile', type=str, required=True, help='Output filename')

    args = parser.parse_args()

    if args.command == 'list':
        PPTXTools.list_slide_layouts(args.template)

    elif args.command == 'merge':
        PPTXMerger.render(args.data, args.template, args.outFile)
        print(f"Wrote {args.outFile}")

    elif args.command == 'clone':
        print(f"Cloned {args.template} --> {args.outFile}")
        PPTXTools.clone_pptx(args.template, args.outFile)
        
    else:
        parser.print_help()

if __name__ == "__main__":
    main()