from distutils.core import setup
import py2exe 

setup(windows=['mealmails.py'],
      options={'py2exe' : { 'bundle_files' : 1,
                            'optimize'     : 2,
                            'compressed'   :  True,
                            'packages'     : ["encodings"],
                            'includes'     : ['encodings']
                          }
      },
      zipfile = None)
