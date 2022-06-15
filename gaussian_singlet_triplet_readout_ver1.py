#!/usr/bin/env python
# coding: utf-8

import sys
import os
from os import chdir
import re
import argparse
import openpyxl
from openpyxl.styles import numbers
import openpyxl.drawing.image
from io import BytesIO
import glob

def get_current_dir():
    if getattr(sys, "frozen", False):
        return path.dirname(sys.executable)
    else:
        return os.getcwd()

def data_extraction(f):
    try:
        with open(f, 'r') as glog:
            s = glog.read()
            
            return s 

    except:
        print("Error for " + f)            

def text_extraction(s): 
    try:
        SS = s .rfind(' Excitation energies and oscillator strengths:')
        SE = s .rfind('SavETr:')
    
        ESS = s[SS:SE]

        ST1 = ESS .find('Singlet')
        ST2 = ESS .rfind('Singlet')

        ESSE = ESS[ST1:ST2]
    
        ST1 = ESSE .find('Singlet')
        ST2 = ESSE .rfind('Singlet')

        ESSE1 = ESSE[ST1:ST2]
    
        Sa1 = ESSE1 .find('      ')
        Se1 = ESSE1 .find(' eV  ')
        Sa2 = ESSE1 .find(' nm  ')
        Se2 = ESSE1 .find(' f=')
        Sa3 = ESSE1 .find('  <S**2>=')
        Se3 = ESSE1 .find('Excited State   ')

        S1E = ESSE1[Sa1:Se1]
        S2E = ESSE1[Se1:Sa2]
        S3E = ESSE1[Se2:Sa3]
        STS = ESSE1[Sa3:Se3]
    
        S1E =  S1E.replace('      ',' ')
        S2E =  S2E.replace(' eV  ',' ')
        S3E =  S3E.replace(' f=',' ')
        STS =  S3E.replace('<S**2>=0.000',' ')

        singlet = str(S1E) + ',' + str(S2E) + ',' + str(S3E)
    
        try:
            TT1 = ESS .find('Triplet')
            TT2 = ESS .rfind('Triplet')

            ETSE = ESS[TT1:TT2]
    
            TT1 = ETSE .find('Triplet')
            TT2 = ETSE .rfind('Triplet')

            ETSE1 = ETSE[TT1:TT2]
    
            Ta1 = ETSE1 .find('      ')
            Te1 = ETSE1 .find(' eV  ')
            Ta2 = ETSE1 .find(' nm  ')

            T1E = ETSE1[Ta1:Te1]
            T2E = ETSE1[Te1:Ta2]

            T1E =  T1E.replace('      ',' ')
            T2E =  T2E.replace(' eV  ',' ')


            triplet =  str(T1E) + ',' + str(T2E)
            sts = singlet + ',' + triplet
            return sts
        
        except:
            triplet = 'error' + ',' + '-' 

            return triplet
            
        return sts
       
    except:
        sts = 'error' + ',' + '-' + ',' + '-' + ',' + '-' + ',' + '-' + ',' + '-'
                
        return sts


def singlet_triplet_extraction(flist):
    output = ''    
    for f in flist:
        s = data_extraction(f)
        sts = text_extraction(s)
        
        f = os.path.basename(f)
        out = str(f.replace('.log','')) + ',' + sts + '\n' 
        output = output + out
            
    return output
            

def main():
    
    dpath = os.path.dirname(sys.argv[0])

    flist = glob.glob(dpath + '/**/*.log', recursive=True)
    print(flist)
    
    chdir(dpath)
    
    output = singlet_triplet_extraction(flist)
           
    s_out = 'name,singlet (eV),singlet (nm),oscillator strength,triplet (eV),triplet (nm)\n'
    final = s_out + output 

    print(final)

    chdir(dpath)
            
    with open("RESULT_TD.csv", "w") as file:
        file.write(final)

if __name__ == "__main__":
    main()        
           
