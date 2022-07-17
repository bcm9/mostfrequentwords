#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Returns n most frequent words in a word document

n = n most frequent words in docfilename in folder

Created on Sun May 10 13:58:13 2020

@author: bcm9
"""
# returns most common words in word document
def mostfrequentwords(folder,docfilename,n):
    # import packages
    import os
    from collections import Counter 
    import readDocx
    
    # pre-processing: set directory, confirm, set filename
    os.chdir(folder)
    os.getcwd() # which working directory
    # doc = docx.Document(docfilename) # load in doc as document
    
    # import doc as string using readDocx
    fulltextstr=(readDocx.getText(docfilename))
      
    # split() returns list of each word within the string
    split_it = fulltextstr.split() 
    # Pass the split_it list to instance of Counter class. 
    wordlist = Counter(split_it) 
      
    # most_common() returns n frequent words 
    most_freq = wordlist.most_common(n) 

    return(most_freq)
