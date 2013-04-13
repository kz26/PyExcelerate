from zipfile import ZipFile, ZIP_DEFLATED

class VirtualDirectory(object):
    def __init__(self, *args, **kwargs):
        self.directories = {}
        self.files = {}

    def __add_directory(self, path):
        subdirs = self.path.split("/")
        if subdirs[0] not in self.directories:
            d = Directory()
            self.directories[subdirs[0]] = d
        if len(subdirs) > 1:
            return self.directories[subdirs[0]].add_directory(subdirs[1:])
        else:
            return self.directories[subdirs[0]]

    def get_or_create_directory(self, path):
        subdirs = self.path.split("/")
        if subdirs[0] not in self.directories:
            d = self.__add_directory(subdirs[0])
        if len(subdirs) > 1:
            return d.get_or_create_directory(subdirs[1:])
        return d

    def add_file(self, path, data):
        ps = path.split("/") 
        if len(ps) == 1:
            self.files[ps[0]] = data
        else:
            d = self.get_or_create_directory(ps[:-1])
            d.files[ps[-1]] = data

    def walk(self):
        dirs = [self] + self.directories.keys()
        while dirs:
            d = dirs.pop()
            yield (d.directories, d.files)
            if d.directories:
                dirs.extend(d.directories)
            

class Filesystem(VirtualDirectory):
    def writeZip(self, path):
        zf = ZipFile(path, 'w', 'ZIP_DEFLATED') 
        for f, d in self.files.items(): 
            zf.writestr(f, d)
