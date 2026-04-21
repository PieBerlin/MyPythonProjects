import argparse
import sys
from classes import Repository

def main():
    parser = argparse.ArgumentParser(description = "A simple Git Clone")

    subparsers = parser.add_subparsers(dest = "command", help = "Available commands")

    # init command
    init_parser = subparsers.add_parser("init", help = "Initialize a new Git repository")

    # add command
    add_parser = subparsers.add_parser("add", help = "Add  files and folders to the staging area ")

    add_parser.add_argument("paths", nargs='+', help="Files and directories to add")

    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return
    
    repo = Repository()

    try :
        #print(args.command)
        if args.command == 'init':   
            if not repo.init():
                print(f"Repository already exists... ")
                return
        elif args.command == "add":
            if not repo.git_dir.exists():
                print(f"Not a git repository")
                return
            
            #print(args.paths)
            
            for path in args.paths:
                repo.add_path(path)

             


        
    except Exception as e:
        print(f"Error; {e}")
        sys.exit(1)



main()