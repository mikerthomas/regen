from datetime import datetime
import xlwt

start_time = datetime.now()

# Ask user for an input regen filename.
regen_file = input("Please type an your REGEN input file name.: ")

# Ask user for an output filename.
output_file = input("Please type output file name ending with .xlsx: ")


with open(regen_file, 'r') as infile: # all data on 1 line
    with open('temp.txt', 'w') as outfile:  # comma delimited file
        data = infile.read()
        # data = data.replace("A001,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,", "")
        # data = data.replace("C001,", "")
        data = data.replace("DIR,DVCFIG", "DIR,,DVCFIG")
        data = data.replace("DVCFIG,COS1", "DVCFIG,,COS1")
        data = data.replace("DIGNODIS,CBKBMAX", "DIGNODIS,,,CBKBMAX")
        data = data.replace("HMUSIC,PMIDX", "HMUSIC,,PMIDX")
        data = data.replace("M2 1 ", "")
        data = data.replace("M1 1 ", "")
        data = data.replace("M2 2 ", "")
        data = data.replace("ADD-", "")
        data = data.replace("CHA-","")
        data = data.replace("CHANGE-", "")
        data = data.replace("COPY-", "")
        data = data.replace("ACT-", "")
        data = data.replace("SET-", "")
        data = data.replace("DEL-", "")

        outfile.write(data)


amo = ["Mike's Tool", 'ACDGP', 'ACDRS_DS', 'ACDRS_RS', 'ACDRT', 'ACDSD', 'ACMSM', 'ACSU', 'ACTDA', 'AFR', 'AGENT', 'ANSU', 'ANUM', 'APC', 
'APESM', 'APESU', 'APP', 'APRT', 'APS', 'ASPIK', 'ASSGN', 'ATCSM', 'AUN', 'AUTH', 'AUTHO', 'BACK', 'BCSM', 'BCSU', 'BDAT', 
'BERUM', 'BERUZ', 'BFDAT', 'BGDAT', 'BPOOL', 'BSSM', 'BSSU', 'BST', 'BUEND', 'BWLST', 'CABA', 'CDBR', 'CDSU', 'CGWB', 'CHEESE', 
'CLIMA', 'CLIST', 'CMP', 'CMUID', 'CODEW', 'COMGR', 'CONSY', 'CONV', 'COP', 'COPY', 'COSSU', 'COT', 'CPCI', 'CPTP', 'CRON', 
'CTIME', 'DAGR', 'DASM', 'DATE', 'DAVF', 'DBC', 'DCIC', 'DCSM', 'DDRSM', 'DDSM', 'DEBUG', 'DEFPP', 'DEFTM', 'DEL', 'DIAGS', 
'DIDCE', 'DIMSU', 'DISPA', 'DISPS', 'DLSM', 'DINT', 'DPSM', 'DSSM', 'DSSU', 'DTIM1', 'DTIM2', 'DTSM', 'DUP', 'DVU', 'FAMOS', 
'FAMUP', 'FBTID', 'FBTN', 'FEACG', 'FEASU', 'FETA', 'FINF', 'FORM', 'FTBL', 'FTCSM', 'FTRNS', 'FUNSU', 'GEFE', 'GENDB', 'GETAB', 
'GETPD', 'GEZAB', 'GEZU', 'GKREG', 'GKTOP', 'GRA', 'GRZW', 'HIDMP', 'HISTA', 'HOTLN', 'INFO', 'INIT', 'JOB', 'KCSU', 'KDEV', 
'KDGZ', 'KNDEF', 'KNFOR', 'KNLCR', 'KNMAT', 'KNPRE', 'KNTOP', 'LANC', 'LAUTH', 'LCSM', 'LDAT', 'LDB', 'LDPLN', 'LDSRT', 'LEMAN', 
'LIN', 'LIST', 'LODR', 'LOGBK', 'LPROF', 'LRNGZ', 'LSCHD', 'LSSM', 'LWCMD', 'LWPAR', 'MFCTA', 'NAVAR', 'PASSW', 'PATCH', 'PERSI', 
'PETRA', 'PRODE', 'PSTAT', 'PTIME', 'RCSU', 'RCUT', 'REFTA', 'REGEN', 'REN', 'REST', 'RICHT', 'RUFUM', 'SA', 'SAVCO', 'SAVE', 
'SBCSU', 'SCREN', 'SCSU', 'SDAT', 'SDSM', 'SELG', 'SELL', 'SELS', 'SIGNL', 'SIPCO', 'SLCB', 'SONUS', 'SPES', 'SSC', 'SSCSU', 
'STMIB', 'SXSU', 'SYNC', 'SYNCA', 'TABTTACSU', 'TAPRO', 'TDCSU', 'TEST', 'TEXT', 'TGER', 'TINFO', 'TLZO', 'TRACA', 'TRACS', 
'TREF', 'TSCSU', 'TSU', 'TTBL', 'TWABE', 'UCSU', 'UPDAT', 'USER', 'USSU', 'VADSM', 'VADSU', 'VBZ', 'VBZA', 'VBZGR', 'VEGAS', 
'VFGKZ', 'VFGR', 'VOICO', 'WABE', 'XAPPL', 'ZAND', 'ZANDE', 'ZAUSL', 'ZIEL', 'ZIELN', 'ZRNGZ']


book = xlwt.Workbook()
#book = xlwt.Workbook(encoding = "utf-8")

for word in amo:
    sheet = book.add_sheet(word)
      

book.save(output_file)

time_elapsed = datetime.now() - start_time 



print('\n\nProgram run time (hh:mm:ss.ms) {}'.format(time_elapsed))

# def print_line():
#     myfile = open('some_file.txt', 'r')
#     for line in myfile:
#         if "251212" in line:
#             print(line)

#Header Rows

# # Puts header row in CSV file.
# with open(delimited_file ,newline='') as f:
#     r = csv.reader(f)
#     data = [line for line in r]
# with open(delimited_file,'w',newline='') as f:
#     w = csv.writer(f)
#     w.writerow(['INDIVIDUAL DEVICE DATA', 'FAMILY', 'EXTEN', 'PHONE FAMILY', 'EXT', 'ASSET-ID KEYBD TYPE', 'SW VERS BOOT-SW', 'TEST HWVERS',
#                  'SW VERSION TDM & HFA', 'SW VERSION TDM & HFA', 'IPADDR', 'REMOVE', 'REMOVE', 'REMOVE', 'DSS VERSION', 'IPADDR', 'REMOVE', 'SW VERS BOOT-SW', 
#                  'SW VERS BOOT-SW', 'TEST HWERS', 'IPADDR', 'IPADDR', 'REMOVE', 'SW VERS BOOT-SW', 'SW VERS BOOT-SW', 'TEST HWERS', 'DPIP CONTROL ADAPTER', 'IPADDR', 'REMOVE', 'SW VERS BOOT-SW', 'TEST HWERS', 
#                  'DPIP CONTROL ADAPTER', 'IPADDR', 'IPADDR', 'SW VERS BOOT-SW', 'TEST HWERS', 'REMOVE', 'REMOVE', 'IPADDR'])
#     w.writerows(data)

'''
with open(delimited_file,'w',newline='') as f:
    w = csv.writer(f)
ACDGRP   w.writerow(['ACDGRP', 'TYPE', 'SEARCH', 'SUPEXT', 'PRIMARY', 'LED', 'ON', 'FLASH', 'WINK'])

ACDRS:DS w.writerow(['ACDRS', 'TYPE', 'DSNUM', 'SHIFT', 'ART', 'EOS', 'ARTDEF'])
ACDRS:RS w.writerow(['ACDRS', 'TYPE', 'RCG', 'SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'])

ACDRT    w.writerow(['ACDRT', 'ART', 'TYPE', 'MAXSTEP / STEP', 'ACT'])
 
ACDDS    w.writerow(['ACDDS', 'TYPE', 'AGTIDLEN_MESSAGE_RCG', 'QFACTOR_CAFDN', 'AGTTARGT_MSGCAT', 'DELAYRBT_MSGNUM', 'SMTIMER', 'SERVCOUT', 'MONPIN', 'AGTFWD', 'NACDAWK'])

ACTDA    w.writerow(['ACTDA', 'TYPE', 'STNO', 'FEATCD'])

AGENT    w.writerow(['AGENT', 'AGTID', 'ACDGRP', 'AGTPOS', 'AGTTYPE', 'AUTOWK', 'PERSANN'])

ANSU     w.writerow(['ANSU', 'TYPE', 'SYSNO', 'SUSY'])

APRT     w.writerow(['APRT', 'TYPE'])

AUN      w.writerow(['AUN', 'GRNO', 'TYPE', 'STNO', 'DISTNO', 'NOTRNG', 'PUSECOND', 'SIGNAL'])

BCSU     w.writerow(['BCSU', 'MTYPE', 'LTG', 'LTU', 'SLOT', 'PARTNO'])

BDAT     w.writerow(['BDAT'])

BERUZ    w.writerow(['BERUZ', 'COSX','DAY', 'COS1T', 'COS1', 'COS2T', 'COS2'])

BFDAT    w.writerow(['BFDAT'])

BUEND    w.writerow(['BUEND', 'TGRP', 'NAME', 'NO', 'TRACENO', 'PRIONO', 'TDDRFLAG', 'GDTRRULE', 'ACDPMGRP', 'CHARCON'])

CGWB     w.writerow(['CGWB', 'MTYPE', 'LTU', 'SLOT', 'TYPE'])

CODEW    w.writerow(['CODEW'])

COMGR    w.writerow(['COMGR'])

COP      w.writerow(['COP', 'COPNO', 'PAR'])

COS      w.writerow(['COSSU', 'TYPE', 'COS'])

COT      w.writerow(['COT', 'COTNO', 'PAR'])

CTIME    w.writerow(['CTIME'])

DCIC     w.writerow(['DCIC'])

DIAAG    w.writerow(['DIAGS', 'PROCID'])

DIDCR    w.writerow(['DIDCR'])

DIMSU    w.writerow(['DIMSU', 'TYPE'])

DNIT     w.writerow(['DNIT', 'DNI', 'INTRTDN', 'ROUTING', 'SARULE', 'DISPLAY', 'ACD', 'TARGET / RCG', 'PRI', 'OVRPRI', 'AUDSRCID', 'THRSHLD', 'REVCHA'])

DTIM1    w.writerow(['DTIM1', 'TYPEDH'])

DTIM2    w.writerow(['DTIM2', 'TYPEDH'])

FEASU    w.writerow(['FEASU', 'TYPE', 'CM'])

FTBL     w.writerow(['FTBL', 'FORMNO', 'TYPE', 'TBL'])

GKREG    w.writerow(['GKREG', 'GWNO', 'GWATTR'])

HOTLN    w.writerow(['HOTLN', 'TYPE', 'HTLNIDX', 'DEST', 'PRECLEVL'])

KCSU     w.writerow(['KCSU', 'TYPE', 'KYNO / PRIMKEY'])

KNFOR    w.writerow(['KNFOR'])

KNDEF    w.writerow(['KNDEF'])

KNLCR    w.writerow(['KNLCR'])

KNMAT    w.writerow(['KNMAT'])

KNPRE    w.writerow(['KNPRE'])

KNTOP    w.writerow(['KNTOP'])

LDAT     w.writerow(['LDAT', 'LROUTE', 'LSVC', 'ALL', 'LVAL', 'TGRP', 'ODR', 'LAUTH', 'CARRIER', 'ZONE'. 'LATTR', 'DNNO', 'VCCYC', 'COTIDX'])

LDPLN    w.writerow(['LDPLN', 'LCRPATT', 'DIPLNUM', 'LDP', 'DPLN', 'LROUT / PROFINDX', 'LAUTH', 'SPC', 'FDSFIELD', 'SDSFIELD', 'PINDP'])

LODR     w.writerow(['LODR', 'ODR', 'CMD / INFO'])

LPROF    w.writerow(['LPROF', 'PROFIDX', 'PROFNAME', 'SRCGRP', 'LRTE'])

LRNGZ    w.writerow(['LRNGZ', 'LST', 'SPDMIN', 'SPDMAX'])

LSCHD    w.writerow(['LSCHD', 'ITR', 'DAY', 'HOUR', 'SCHED', 'MINUTE'])

LWPAR    w.writerow(['LWPAR'])

MFCTA    w.writerow(['MFCTA'])

NAVAR    w.writerow(['NAVAR', 'NOPTNO', 'TYPE', 'STNO / CD'])

PRODE    w.writerow(['PRODE', 'TYPE', 'PVCDNO'])

PSTAT    w.writerow(['PSTAT', 'TYPE', 'PERIOD / ERROR', 'MODE / MAXCOUNT / SWITCH / CPTYPE', 'SWITCH'])

PTIME    w.writerow(['PTIME'])

RICHT    w.writerow(['RICHT'])

RCSU     w.writerow(['RCSU', 'PEN', 'NO', 'ACT', 'OPMODE', 'OPTYPE', 'COFIDX', 'CLASSMRK', 'RECAPL', 'RECNUM', 'LENGTH'])

REFTA    w.writerow(['REFTA', 'TYPE', 'PEN', 'PRI', 'BLOCK', 'READYASY'])

RUFUM    w.writerow(['RUFUM'])

SA       w.writerow(['SA', 'TYPE', 'CD', 'ITR', 'STNO'])

SBCSU    w.writerow(['SBCSU', 'STNO', 'OPT', 'CONN', 'PEN', 'DVCFIG', 'TSI', 'COS1', 'COS2', 'LCOSV1', 'LCOSV2', 'LCOSD1', 'LCOSD2', 'DPLN', 'ITR', 'SSTNO', 'COSX', 'SPDI', 'SPDC1', 'IDCR', 'REP', 'STD', 'SECR', 'INS', 'ALARMNO', 'RCBKNA', 'DSSTNB', 'DIGNODIS', 'OPTICA', 'OPTIDA', 'CBKBMAX', 'HEADSET', 'HSKEY', 'CBKNAMB', 'TEXTSEL', 'HMUSIC', 'CALLOG', 'PMIDX', 'COMGRP'])

SCREN    w.writerow(['SCREN'])

SCSU     w.writerow(['SCSU', 'STNO', 'PEN', 'DVCFIG', 'DPLN', 'ITR', 'COS1', 'COS2', 'LCOSV1', 'LCOSV2', 'LCOSD1', 'LCOSD2', 'SPDC1', 'SPDC2', 'COSX' 'SPDI', 'COFIDX', 'CCTIDX',

PERSI    w.writerow(['PERSI'])

SDAT     w.writerow(['SDAT', 'STNO', 'TYPE'])

SELG     w.writerow(['SELG'])

SELS     w.writerow(['SELS'])

SIPCO    w.writerow(['SIPCO', 'TYPE'])

SSC      w.writerow(['SSC'])

SSCSU    w.writerow(['SSCSU'])

STMIB    w.writerow(['STMIB', 'MTYPE', 'LTU', 'TYPE'])

TACSU    w.writerow(['TACSU', 'PEN', 'COTNO', 'COPNO', 'DPLN', 'ITR', 'COS', 'LCOSV', 'LCOSD', 'TGRP', 'COFIDX', 'CCT', 'DESTNO', 'NNO', 'ALARMNO', 'CARRIER', 'ZONE', 'LIN', 'CIDDGTS', 'CBMATTR', 'SCRGRP', 'CLASSMARK', 'CTCCID', 'DITIDX', 'TRTBL', 'RULEIDX', 'ATNTYP', 'HMUSIC', 'INS', 'DEVTYPE', 'DEV', 'MFCVAR', 'SUPPRESS', 'DGTCNT', 'TESTNO', 'CIRCIDX', 'CDRINT', 'CCTINFO', 'DIALTYPE', 'DIALVAR', 'COEX'])

TAPRO    w.writerow(['TAPRO'])

TDCSU    w.writerow(['TDCSU', 'OPT', 'PEN', 'COTNO', 'COPNO', 'DPLN', 'ITR', 'COS', 'LCOSV', 'LCOSD', 'CCT', 'DESTNO', 'PROTVAR', 'SEGMENT', 'DEDSVC', 'TRTBL', 'SIDANI', 'ATNTYP', 'CBMATTR', 'TCHARG', 'SUPPRESS', 'TRACOUNT', 'SATCOUNT', 'NNO', 'ALARMNO', 'FIDX', 'CARRIER', 'ZONE', 'COTX', 'FWDX', 'CHIMAP', 'UUSCCX', 'UUSCCY', 'FNIDX', 'NWMUXTIM', 'SRCGRP', 'CLASSMARK', 'TCCID', 'TGRP', 'SRCHMODE', 'INS', 'SECLEVEL', 'HMUSIC', 'CALLTIM', 'WARNTIM', 'DEV', 'BCHAN', 'BCNEG', 'BCGR', 'LWPP', 'LWLT', 'LWPS', 'LWR1', 'LWR2', 'DMCALLWD'])       
 
TGER     w.writerow(['TGER'])

TSCSU    w.writerow(['TSCSU'])

TWABE    w.writerow(['TWABE'])

UCSU     w.writerow(['UCSU', 'UNIT', 'LTG', 'LTU', 'LTPARTNO'])

VADSU    w.writerow(['VADSU'])

VBZ      w.writerow(['VBZ'])

VBZA     w.writerow(['VBZA'])

VBZGR    w.writerow(['VBZGR'])

VFGKZ    w.writerow(['VFGKZ', 'TYPE', 'CD', 'CPS', 'ATNDGR'])

VFGR     w.writerow(['VFGR', 'ATNDGR'])

VOICO    w.writerow(['VOICO'])

WABE     w.writerow(['WABE', 'CD', 'DAR', 'CHECK'])

ZAND     w.writerow(['ZAND', 'TYPE'])

ZANDE    w.writerow(['ZANDE', 'TYPE'])

ZIEL     w.writerow(['ZIEL', 'TYPE', 'SRCNO', 'SI', 'DESTNOF'])

ZIELN    w.writerow(['ZIELN', 'STNO', 'TYPE', 'KYNO', 'LEVEL'])
'''






