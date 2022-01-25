import subprocess
sasPath = '//sas-vm/DEV2021/Data/ADAM/programs/adae.sas'
logPath = '//sas-vm/DEV2021/Data/ADAM/programs/adae.log'
outPath = '//sas-vm/DEV2021/Data/ADAM/programs/adae.lst'
print(f"C:/Program Files/SASHome/SASFoundation/9.4/sas.exe -sysIn \"{sasPath}\" -log \"{logPath}\" -print \"{outPath}\"")
# subprocess.call(f"C:/Program Files/SASHome/SASFoundation/9.4/sas.exe -sysIn \"{sasPath}\" -log \"{logPath}\" -print \"{outPath}\"")