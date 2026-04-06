import os
import comtypes.client
from pptxtopdf import convert

single_or_multi = input("Files to convert [s] or [m]: ")
single_or_multi = single_or_multi.lower()

if single_or_multi == 's':
    input_dir = input("Enter input directory path: ")
    output_dir = input("Enter output directory path: ")
    convert(input_dir,output_dir)

elif single_or_multi == 'm':
    s_or_m_output = input("Different or same output directory? [s] or [d] ")
    s_or_m_output = s_or_m_output.lower()
    i = 1
    inp_dir_list = []
    out_dir_list = []
    
    if s_or_m_output == 's':
        while True:
            item = input(f"Enter input directory path {i}, type stop to stop: ")
            i += 1
            
            if item.lower() != 'stop': 
                inp_dir_list.append(item)
            elif item.lower() == 'stop':
                output_dir = input("Enter output directory path: ")
                break

        for i in range(len(inp_dir_list)):
            convert(inp_dir_list[i],output_dir)
    
    if s_or_m_output == 'd':
        while True:
            item = input(f"Enter input directory path {i}, type stop to stop: ")
            
            if item.lower() != 'stop': 
                output_item = input(f"Enter output directory path {i}: ")
                inp_dir_list.append(item)
                out_dir_list.append(output_item)
            elif item.lower() == 'stop':
                break
            i += 1
            
        for i in range(len(inp_dir_list)):
            convert(inp_dir_list[i],out_dir_list[i])
