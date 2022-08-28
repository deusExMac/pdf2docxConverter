
#----------------------------------------------------------
#
#
# This script converts pdf files into ms word .docx files.
# The script collects files matching a pattern inside
# a directory and converts matching pdf files into .docx
# (keeping filenames)
#
# See usage for a list of command line options.
#
# v0.2@rd20082022
#----------------------------------------------------------


import os
import sys
import re
import time
import pathlib

import argparse
from pdf2docx import Converter

# The following two classes are used to parse
# arguments on the shell 'scommand line
#
# They specialise ArgumentParser exceptions so that
# they throw exceptions in case an error is encounterred
class ArgumentParserError(Exception): pass
  
class ThrowingArgumentParser(argparse.ArgumentParser):
      def error(self, message):
          raise ArgumentParserError(message)




# Return a list of filenames in targetDir
# that matches regular expression pattern.
# Fullpaths of matching filenames are returned in
# a list
def listDirectoryFiles(targetDir, pattern):

    cwDir = os.getcwd()
    
    os.chdir(targetDir)
    files = filter(os.path.isfile, os.listdir(targetDir)) 
    files = [os.path.join(targetDir, f) for f in files if re.search(pattern, f) is not None] # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    os.chdir(cwDir)
    return(files)




# sAmount in seconds
def countdown(sAmount):
    t = int(sAmount)
    while t:
        mins, secs = divmod(t, 60)
        timer = '{:02}:{:02}'.format(mins, secs)
        print(f'{timer} ', end="")
        time.sleep(1)
        t -= 1
        
    print('\nStarting...\n')





def main():

   # Parse command line arguments   
   cmdLineArgs = ThrowingArgumentParser()
   cmdLineArgs.add_argument('directory', nargs=argparse.REMAINDER, default='./' )
   cmdLineArgs.add_argument('-p', '--pattern', type=str, default='\.pdf$' )
   cmdLineArgs.add_argument('-s', '--start', type=int, nargs='?', default=1 )
   cmdLineArgs.add_argument('-e', '--end', type=int, nargs='?', default=None )
   cmdLineArgs.add_argument('-o', '--outputdir', type=str, nargs='?', default='./' )
   
   cmdLineArgs.add_argument('-P', '--password', type=str, nargs='?', default=None )

   cmdLineArgs.add_argument('-N',  '--nodelete', action='store_true')
   cmdLineArgs.add_argument('-G',  '--debug', action='store_true')
   
   args = vars( cmdLineArgs.parse_args() )



   #
   # Post processing some command line arguments.
   #
   
   if len(args['directory']) == 0:
      args['directory'] = './'
   else:
       #print('Checking is ', args['directory'], 'relative')
       #print('is relative:', os.path.isabs(args['directory'][0]))
       if not os.path.isabs(args['directory'][0]):
          args['directory'][0] = pathlib.Path(__file__).parent / args['directory'][0]
          #print(args['directory'][0] )



   # On the command line, start page is given in human counting i.e.
   # first page starts at 1. Here, this is transformed in python counting (zero-based)
   if args['start'] <= 0:
      print('Invalid starting page ', args['start'], '. At the command line, page counting starts at 1. Use 1,2,3,...')
      return(-3)
    
   # make zero-based
   args['start'] = args['start'] - 1


   if args['debug']: 
      print(args)

    
   
   # Get list of matching files
   matchingList = listDirectoryFiles(args['directory'][0], args['pattern'])
   if not matchingList:
      print('No filenames matching regular expression [', args['pattern'], '] in directory [', args['directory'][0], ']', sep='')
      print('\nUsage: p2dConverter [-p pattern="\.pdf$"] [-s frompagenumber=1] [-e topagenumber=None] [-o outputdir="./"] [-P password=None] [-N] [-G] [source directory="./"]')
      return(-1)

   #print('\nUsage: p2dConverter [-p pattern="\.pdf$"] [-s frompagenumber=1] [-e topagenumber=None] [-o outputdir="./"] [-P password=None] [-N] [-G] [source directory="./"] \n') 
   print('\nConverting to .docx all .pdf with following parameters:')
   print('\tSource directory: [', args['directory'], ']')
   print('\tFilename pattern: [', args['pattern'], '](Total of ', len(matchingList), 'filenames match pattern)')
   print('\tOutput directory: [', args['outputdir'], ']')
  

   try:

    # Initializing some variables.
    # Doing it at this point as some of these may be used during exceptions.
    nConverted = 0
    
    if not args['nodelete']:  
      print('\n\tSettings indicate that if destination files already exists, these will be deleted. Use -N option to avoid overwriting existing files.')
      print('\tYou have 5 seconds to think and terminate the process.')

    print('\nContinue? (Type Control-C to terminate.)') 
    time.sleep( 3.9 )
    countdown(5)

    
    #
    #
    # Iterate ove collected files matching pattern and start converting
    # them one by one.
    #
    #
    
    for n, sourceFilename in enumerate(matchingList): 
         path, filename = os.path.split(sourceFilename)
         if args['outputdir'] != '':
            path = args['outputdir']
            
         destinationFilename = path + '/' + os.path.splitext(filename)[0] + '.docx'
         print(n+1, ') Converting [', sourceFilename, ']...', sep='')

         if os.path.isfile(destinationFilename):
            if args['nodelete']:
               print('\tDestination file [', destinationFilename, '] already exists. Skipping file due to -N option. Nothing converted.', sep='') 
               continue
            else:
               print('\tDestination file exists. Deleting', destinationFilename, '...', end='') 
               os.remove( destinationFilename )
               print('Dοne.')

         cv = Converter(sourceFilename, password=args['password'])
         try:   
           cv.convert(destinationFilename, start=args['start'], end=args['end'])
           cv.close()
           nConverted += 1
         except Exception as cvException:
           print('Conversion error.', str(cvException))  
           cv.close() # TODO: Is this the right place to free resources?
           continue
     
   except KeyboardInterrupt:  
       print('Control-C seen. Terminating.', sep='')
       print('Converted a total of', nConverted, 'out of', len(matchingList), 'files')
       #cv.close()  
       return(-2)
   
   print('\nDone processing', nConverted, 'out of', len(matchingList), 'files')
   return(0)





# main guard
if __name__ == '__main__':
   main()

   
