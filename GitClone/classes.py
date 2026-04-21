from __future__ import annotations 
from pathlib import Path
import json 
import hashlib
from typing import Dict
import zlib 


class GitObject:
    def __init__(self,obj_type:str , content:bytes):
        self.type = obj_type 
        self.content = content 
    
    def hash(self) -> str:
        # (<type> <size> \0 <content>)
        header = f"{self.type} {len(self.content)}\0".encode()

        return hashlib.sha1(header + self.content).hexdigest()

    def serialize(self) -> bytes:
        header = f"{self.type} {len(self.content)}\0".encode()
        return zlib.compress(header + self.content)
    
    @classmethod
    def deserialize(cls, data: bytes)-> GitObject:
        decompressed = zlib.decompress(data)
        null_idx = decompressed.find(b"\0")
        header = decompressed[:null_idx]
        content = decompressed[null_idx + 1:]

        obj_type, size = header.split(" ") # but we dont use the size for now

        return cls(obj_type,content)

# stores only the file content     
class Blob(GitObject):
    def __init__(self, content:bytes):
        super().__init__("blob", content)

    def get_content(self) -> bytes:
        return self.content
    




class Repository:
    def __init__(self,path="."):
        self.path = Path(path).resolve() # pass the current  absolute path 
        self.git_dir = self.path / ".pygit" # use .pytgit instead of .git to avoid conflicts

        # .git/objects
        self.objects_dir = self.git_dir / " objects"

        # .git/refs
        self.refs_dir = self.git_dir / " refs" 
        self.heads_dir = self.refs_dir / "heads"

        # HEAD file
        self.head_file = self.git_dir / "HEAD"

        # .git/index -- index area 
        self.index_file = self.git_dir / "index"

    def init(self) -> bool:

        if self.git_dir.exists():
            return False

        # create directories here
        self.git_dir.mkdir()
        self.objects_dir.mkdir()
        self.refs_dir.mkdir()
        self.heads_dir.mkdir()
        
        #create initial HEAD pointing to a branch 
        self.head_file.write_text("ref: refs/heads/master\n")

        self.save_index({})

        print(f"Initialized empty Git repository in {self.git_dir}")
        return True
    
    # we use GitObject so that this function can be reused
    def store_object(self,obj:GitObject):
        obj_hash = obj.hash()

        # create a directory with the first two characters of the hash 
        obj_dir = self.objects_dir / obj_hash[:2]
        obj_file = obj_dir / obj_hash[2:]

        if not obj_file.exists():
            
            obj_dir.mkdir(exist_ok=True)
            
            obj_file.write_bytes(obj.serialize())
            

        return obj_hash



    def load_index(self) -> Dict[str,str]:
        if not self.index_file.exists():
            return {} 
        try :
            return json.loads(self.index_file.read_text())

        except:
            return {}

    def save_index(self,index:Dict[str,str]):
        self.index_file.write_text(json.dumps(index,indent =2))

    
    # helper functions
    def add_file(self,path:str):
        full_path = self.path / path  
        if not full_path.exists():  
            raise FileNotFoundError(f"Path {path} not found")
        
        # read the file contents
        content = full_path.read_bytes()
        
        # create a BLOB object from the content
        blob = Blob(content)
        
        #store the blob onject in the database (.git /objects)
        blob_hash = self.store_object(blob)
        
        #update index file to include this file
        index = self.load_index()

        index[path] = blob_hash 

        # save the index file back 
        self.save_index(index)

        print(f"Added {path}")



    def add_directory(self,path):
        full_path = self.path / path  
        if not full_path.exists():
            raise FileNotFoundError(f"Directory {path} not found")
        if not full_path.is_dir():
            raise ValueError(f"{path} is not a directory")

        

        # recursively traverse the directory and add all files and subdirectories
        index = self.load_index()
        added_count = 0
        for file_path in full_path.rglob("*"):
            if file_path.is_file():
                if ".pygit" in file_path.parts:
                    continue
                if ".git" in file_path.parts:
                    continue
                # create blob object
                content = file_path.read_bytes()
                blob = Blob(content)
                blob_hash = self.store_object(blob)
                #update index
                relative_path =str(file_path.relative_to(self.path))
                index[relative_path] = blob_hash
                added_count += 1

        self.save_index(index) 
        if added_count > 0 :
            print(f"Added {added_count} files from directory {path}")
        else:
            print(F"Directory {path} already upto date ")
        #create blob objects for each file and store them in the database

        # update the index to include all the files 

    
    def add_path(self,path:str)-> None:
        full_path = self.path / path 

        if not full_path.exists():
            raise FileNotFoundError(f"Path {path} not found")
        
        if full_path.is_file():
            self.add_file(path)
        elif full_path.is_dir():
            self.add_directory(path)
        else :
            raise ValueError(f"{path} is neither a file nor a directory")
        
    







        # create 








