#!/usr/bin/env python
# coding: utf-8

import os
from os import chdir
import sys
import argparse
import openpyxl
from openpyxl.styles import numbers
import openpyxl.drawing.image
from io import BytesIO
import glob

def data_extraction(f):
    try:
        with open(f, 'r') as glog:
            s = glog.read()
            
            return s 

    except:
        print("Error for " + f)        


def text_extraction(s): 
    try:
        E1 = s .rfind(' SCF Done:')
        E2 = s .rfind(' A.U. after')
        P1 = s .rfind(' Occupied')
        P2 = s .rfind(' Virtual')
        P3 = s .rfind(' The electronic state is')
        P4 = s .rfind(' Condensed to atoms (all electrons):')

        tot_ene = s[E1:E2]
        occ_num = s[P1:P2]
        virt_num = s[P2:P3]
        ene = s[P3:P4]
        
        E3 = tot_ene .rfind('-')
        P5 = ene .rfind(' The electronic state is')
        P6 = ene .find(' Alpha virt. eigenvalues --')
        P7 = ene .rfind(' Condensed to atoms (all electrons):')

        tot_ene = tot_ene[E3:]        
        occ_ene = ene[P5:P6]
        virt_ene = ene[P6:P7]

        tot_ene = tot_ene.replace(' ','')
        occ_num = occ_num.count(' (')
        occ_num = int(occ_num)
        virt_num = virt_num.count(' (')
        virt_num = int(virt_num)
        occ_ene = occ_ene.replace('The electronic state is 1-A.', '')
        occ_ene = occ_ene.replace('Alpha  occ. eigenvalues --  ','')
        occ_ene = occ_ene.replace('¥n ',' ')
        occ_ene = occ_ene.split()
        HOMO = occ_ene[occ_num - 1]
        virt_ene = virt_ene.replace('Alpha virt. eigenvalues --','')
        virt_ene = virt_ene.replace('¥n ',' ')
        virt_ene = virt_ene.split()
        LUMO = virt_ene[0]
                            
        au_ev = 27.211

        HOMO_eV = float(HOMO) *  au_ev
        LUMO_eV =  float(LUMO) *  au_ev
        
        tot_HL = str(tot_ene) + ',' + str(HOMO) + ',' + str(HOMO_eV) + ',' + str(LUMO) + ',' + str(LUMO_eV)
        
        return tot_HL
            
    except:
        tot_HL = 'error' + ',' + 'error' + ',' + 'error'
                
        return tot_HL 


def HOMO_LUMO_extraction(flist):
    output = ''    
    for f in flist:
        s = data_extraction(f)
        tot_HL = text_extraction(s)
            
        out = str(f.replace('.log','')) + ',' + tot_HL + '\n' 
        output =output +out
            
    return output
            
    
def main():
    
    dpath = os.path.dirname(sys.argv[0])

    flist = [os.path.basename(p) for p in glob.glob(dpath + '/*.log', recursive=True) if os.path.isfile(p)]
    print(flist)
    
    chdir(dpath)
    
    output = HOMO_LUMO_extraction(flist)

    s_out = 'name,E (A.U.),HOMO (A.U.),HOMO (eV),LUMO (A.U.),LUMO (eV)\n'
    final = s_out + output

    print(final)

    chdir(dpath)
            
    with open("RESULT.csv", "w") as file:
        file.write(final)

if __name__ == "__main__":
    main()        

