"""
Credits to delimitry(https://github.com/delimitry/compressed_rtf) for RTF Decompression method(w/ CRC32 calculation)
MSG file format reference:
http://www.fileformat.info/format/outlookmsg/index.htm

"""
import email
import re
import os
import olefile
from argparse import ArgumentParser
import struct
from cStringIO import StringIO

urlList = []
attachmentsData = []
def main():
    parser = ArgumentParser(description='Extract attachments and URLs from mail file (MSG | MIME).')
    parser.add_argument('file',help='Path to mail file')
    args = parser.parse_args()
    msgFile = args.file
    if olefile.isOleFile(msgFile):
        extractFromMSG(msgFile)
    else:
        isMime = False
        with open(msgFile,'r') as openedFile:
            line = openedFile.read()
            mimePattern = re.search('Content-Type: (multipart|image|application|audio|message|text|x-token|video)\/',line)
            if mimePattern:
                isMime = True
                openedFile.close()
        if isMime:
            extractFromMIME(msgFile)
        else:
            print("Not an e-mail file!")
            exit()
    printURLs()
    createDir(msgFile)
    saveAttachments(attachmentsData)
def createDir(msgFile):
    global path
    path = os.path.join(os.getcwd(),'attachments')
    path = os.path.join(path,os.path.basename(msgFile))
    if not os.path.exists(path):
        os.makedirs(path)
def extractFromMIME(msgFile):
        with open(msgFile) as emailFile:
              msg = email.message_from_file(emailFile)
        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                attachmentData = part.get_payload(decode=True)
                attachmentsData.append((attachmentData,filename))
            else:
                if 'text' in part.get_content_type():
                    text = part.get_payload(decode=True)
                    extractURLRegEx(text)
                    extractFromHREF(text,len(text))
def extractFromMSG(msgFile):
    #Trails with 102 mean that the stream is binary information, 1F is unicode, 1E is ASCII
    msg = olefile.OleFileIO(msgFile)
    attachDirs = []
    for dir in msg.listdir():
        if dir[0].startswith('__attach') and dir[0] not in attachDirs: # Attachments are stored in __attach_version1.0_#000000? folders where ? is the number of the respective attachment
            attachDirs.append(dir[0])
        if ('__substg1.0_1000001F') == dir[0]: #Message body is stored in __substg1.0_1000001F
            msgBodyFilename = dir[0]
            hasBody = True
        if ('__substg1.0_10090102') == dir[0]: #Href tags are stored in __substg1.0_10090102
            rtfCompressedBinary = dir[0]
            hasBodyInfo = True
    if hasBody and hasBodyInfo:
        messageBody = msg.openstream(msgBodyFilename).read()
        messageBody = messageBody.translate(None,'\0')
        rtfCompressedBinary = msg.openstream(rtfCompressedBinary).read()
        extractURLRegEx(messageBody)
        decompressedRTF = decompress(rtfCompressedBinary)
        extractFromHREF(decompressedRTF,len(decompressedRTF))
        for dir in attachDirs:
            attachmentData = msg.openstream(dir+'/'+'__substg1.0_37010102').read() # Attachment Data is stored in __substg1.0_37010102 stream of each attachment storage
            attachmentFilename = msg.openstream(dir+'/'+'__substg1.0_3707001F').read() # Long-filenames are stored in __substg1.0_3707001F stream of each attachment storage
            attachmentsData.append((attachmentData,attachmentFilename))
        msg.close()
    else:
        print("Not an e-mail file!")
        exit()
def saveAttachments(attachmentsData):
    if attachmentsData:
        print('\nList of Attachments:\n**********************')
        for attachmentData in attachmentsData:
            data, filename = attachmentData
            filepath = os.path.join(path,filename.translate(None,'\0')) #Translate to remove any null byte from filename
            with open(filepath, 'wb') as fp:
                fp.write(data)        
            print('Attachment: ' + filename + ' has been extracted and stored in ' + path)
        print('**********************')
    else:
        print("No attachments found.")
def extractURLRegEx(text):
        global urlList
        regExURLs= re.findall(ur'(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:\'".,<>?\xab\xbb\u201c\u201d\u2018\u2019]))',text)
        for url in regExURLs:
            urlList.append(''.join(url))
def extractFromHREF(text,textSize):
    global urlList
    start = 0
    while start < textSize:
        value = text.find('href',start,textSize)
        if value == -1:
            break
        else:
            endofRef = text.find('>',value,textSize)
            start = endofRef
            url = re.search(ur'href\s*=\s*(.*)',text[value:endofRef])
            if url:
                url = url.group(0)
                url = url[url.find('=',0,len(url))+1:len(url)]
                url = formatURL(url)
                if not url.startswith(('fax:','tel:','mailto:','cid:')) and not (url.startswith('~~') and url.endswith('~~')) and not (url.startswith('#') and url[4] == 'A'):
                    urlList.append(url)
def formatURL(url):
    url.strip() #remove leading and trailing whitespaces
    if '"' in url[0] and  '"' in url[len(url) - 1]: #remove quotation marks around the URL
        url = url[1:len(url)-1]
    return url
def printURLs():
    global urlList
    list(set(urlList)) #Remove duplicates
    if not urlList:
        print('No URLs found.')
    else:
        print('List of extracted URLs:\n**********************')
        print('\n'.join(urlList))
        print('**********************')

"""
def beautifyURLs():
    global urlList
    for idx,url in enumerate(urlList):
        if not url.startswith(('https','http')) # to remove any redundant
            urlList[idx] = formatURL(url)
    return urlList
"""
table = [
    0x00000000, 0x77073096, 0xee0e612c, 0x990951ba,
    0x076dc419, 0x706af48f, 0xe963a535, 0x9e6495a3,
    0x0edb8832, 0x79dcb8a4, 0xe0d5e91e, 0x97d2d988,
    0x09b64c2b, 0x7eb17cbd, 0xe7b82d07, 0x90bf1d91,
    0x1db71064, 0x6ab020f2, 0xf3b97148, 0x84be41de,
    0x1adad47d, 0x6ddde4eb, 0xf4d4b551, 0x83d385c7,
    0x136c9856, 0x646ba8c0, 0xfd62f97a, 0x8a65c9ec,
    0x14015c4f, 0x63066cd9, 0xfa0f3d63, 0x8d080df5,
    0x3b6e20c8, 0x4c69105e, 0xd56041e4, 0xa2677172,
    0x3c03e4d1, 0x4b04d447, 0xd20d85fd, 0xa50ab56b,
    0x35b5a8fa, 0x42b2986c, 0xdbbbc9d6, 0xacbcf940,
    0x32d86ce3, 0x45df5c75, 0xdcd60dcf, 0xabd13d59,
    0x26d930ac, 0x51de003a, 0xc8d75180, 0xbfd06116,
    0x21b4f4b5, 0x56b3c423, 0xcfba9599, 0xb8bda50f,
    0x2802b89e, 0x5f058808, 0xc60cd9b2, 0xb10be924,
    0x2f6f7c87, 0x58684c11, 0xc1611dab, 0xb6662d3d,
    0x76dc4190, 0x01db7106, 0x98d220bc, 0xefd5102a,
    0x71b18589, 0x06b6b51f, 0x9fbfe4a5, 0xe8b8d433,
    0x7807c9a2, 0x0f00f934, 0x9609a88e, 0xe10e9818,
    0x7f6a0dbb, 0x086d3d2d, 0x91646c97, 0xe6635c01,
    0x6b6b51f4, 0x1c6c6162, 0x856530d8, 0xf262004e,
    0x6c0695ed, 0x1b01a57b, 0x8208f4c1, 0xf50fc457,
    0x65b0d9c6, 0x12b7e950, 0x8bbeb8ea, 0xfcb9887c,
    0x62dd1ddf, 0x15da2d49, 0x8cd37cf3, 0xfbd44c65,
    0x4db26158, 0x3ab551ce, 0xa3bc0074, 0xd4bb30e2,
    0x4adfa541, 0x3dd895d7, 0xa4d1c46d, 0xd3d6f4fb,
    0x4369e96a, 0x346ed9fc, 0xad678846, 0xda60b8d0,
    0x44042d73, 0x33031de5, 0xaa0a4c5f, 0xdd0d7cc9,
    0x5005713c, 0x270241aa, 0xbe0b1010, 0xc90c2086,
    0x5768b525, 0x206f85b3, 0xb966d409, 0xce61e49f,
    0x5edef90e, 0x29d9c998, 0xb0d09822, 0xc7d7a8b4,
    0x59b33d17, 0x2eb40d81, 0xb7bd5c3b, 0xc0ba6cad,
    0xedb88320, 0x9abfb3b6, 0x03b6e20c, 0x74b1d29a,
    0xead54739, 0x9dd277af, 0x04db2615, 0x73dc1683,
    0xe3630b12, 0x94643b84, 0x0d6d6a3e, 0x7a6a5aa8,
    0xe40ecf0b, 0x9309ff9d, 0x0a00ae27, 0x7d079eb1,
    0xf00f9344, 0x8708a3d2, 0x1e01f268, 0x6906c2fe,
    0xf762575d, 0x806567cb, 0x196c3671, 0x6e6b06e7,
    0xfed41b76, 0x89d32be0, 0x10da7a5a, 0x67dd4acc,
    0xf9b9df6f, 0x8ebeeff9, 0x17b7be43, 0x60b08ed5,
    0xd6d6a3e8, 0xa1d1937e, 0x38d8c2c4, 0x4fdff252,
    0xd1bb67f1, 0xa6bc5767, 0x3fb506dd, 0x48b2364b,
    0xd80d2bda, 0xaf0a1b4c, 0x36034af6, 0x41047a60,
    0xdf60efc3, 0xa867df55, 0x316e8eef, 0x4669be79,
    0xcb61b38c, 0xbc66831a, 0x256fd2a0, 0x5268e236,
    0xcc0c7795, 0xbb0b4703, 0x220216b9, 0x5505262f,
    0xc5ba3bbe, 0xb2bd0b28, 0x2bb45a92, 0x5cb36a04,
    0xc2d7ffa7, 0xb5d0cf31, 0x2cd99e8b, 0x5bdeae1d,
    0x9b64c2b0, 0xec63f226, 0x756aa39c, 0x026d930a,
    0x9c0906a9, 0xeb0e363f, 0x72076785, 0x05005713,
    0x95bf4a82, 0xe2b87a14, 0x7bb12bae, 0x0cb61b38,
    0x92d28e9b, 0xe5d5be0d, 0x7cdcefb7, 0x0bdbdf21,
    0x86d3d2d4, 0xf1d4e242, 0x68ddb3f8, 0x1fda836e,
    0x81be16cd, 0xf6b9265b, 0x6fb077e1, 0x18b74777,
    0x88085ae6, 0xff0f6a70, 0x66063bca, 0x11010b5c,
    0x8f659eff, 0xf862ae69, 0x616bffd3, 0x166ccf45,
    0xa00ae278, 0xd70dd2ee, 0x4e048354, 0x3903b3c2,
    0xa7672661, 0xd06016f7, 0x4969474d, 0x3e6e77db,
    0xaed16a4a, 0xd9d65adc, 0x40df0b66, 0x37d83bf0,
    0xa9bcae53, 0xdebb9ec5, 0x47b2cf7f, 0x30b5ffe9,
    0xbdbdf21c, 0xcabac28a, 0x53b39330, 0x24b4a3a6,
    0xbad03605, 0xcdd70693, 0x54de5729, 0x23d967bf,
    0xb3667a2e, 0xc4614ab8, 0x5d681b02, 0x2a6f2b94,
    0xb40bbe37, 0xc30c8ea1, 0x5a05df1b, 0x2d02ef8d
]


def crc32(data):
    """
    Calculate CRC32 from given data bytes
    """
    stream = StringIO(data)
    crc_value = 0x00000000
    while True:
        char = stream.read(1)
        if not char:
            break
        table_pos = (crc_value ^ ord(char)) & 0xff
        intermediate_value = crc_value >> 8
        crc_value = table[table_pos] ^ intermediate_value
    return crc_value

INIT_DICT = (
    '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\'
    'fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New '
    'RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\'
    'f0\\fs20\\b\\i\\u\\tab\\tx'
)

INIT_DICT_SIZE = 207
MAX_DICT_SIZE = 4096

COMPRESSED = 'LZFu'
UNCOMPRESSED = 'MELA'


def decompress(data):
    """
    Decompress data
    """
    # set init dict
    init_dict = list(INIT_DICT)
    init_dict += ' ' * (MAX_DICT_SIZE - INIT_DICT_SIZE)
    if len(data) < 16:
        raise Exception('Data must be at least 16 bytes long')
    write_offset = INIT_DICT_SIZE
    output_buffer = ''
    # make stream
    in_stream = StringIO(data)
    # read compressed RTF header
    comp_size = struct.unpack('<I', in_stream.read(4))[0]
    raw_size = struct.unpack('<I', in_stream.read(4))[0]
    comp_type = in_stream.read(4)
    crc_value = struct.unpack('<I', in_stream.read(4))[0]
    # get only data
    contents = StringIO(in_stream.read(comp_size - 12))
    if comp_type == COMPRESSED:
        # check CRC
        if crc_value != crc32(contents.read()):
            raise Exception('CRC is invalid! The file is corrupt!')
        contents.seek(0)
        end = False
        while not end:
            val = contents.read(1)
            if not val:
                break
            control = '{0:08b}'.format(ord(val))
            # check bits from LSB to MSB
            for i in xrange(1, 9):
                if control[-i] == '1':
                    # token is reference (16 bit)
                    val = contents.read(2)
                    if not val:
                        break
                    token = struct.unpack('>H', val)[0]  # big-endian
                    # extract [12 bit offset][4 bit length]
                    offset = (token >> 4) & 0b111111111111
                    length = token & 0b1111
                    # end indicator
                    if write_offset == offset:
                        end = True
                        break
                    actual_length = length + 2
                    for step in xrange(actual_length):
                        read_offset = (offset + step) % MAX_DICT_SIZE
                        char = init_dict[read_offset]
                        output_buffer += char
                        init_dict[write_offset] = char
                        write_offset = (write_offset + 1) % MAX_DICT_SIZE
                else:
                    # token is literal (8 bit)
                    val = contents.read(1)
                    if not val:
                        break
                    output_buffer += val
                    init_dict[write_offset] = val
                    write_offset = (write_offset + 1) % MAX_DICT_SIZE
    elif comp_type == UNCOMPRESSED:
        return contents.read(raw_size)
    else:
        raise Exception('Unknown type of RTF compression!')
    return output_buffer
if __name__ == '__main__':
    main()