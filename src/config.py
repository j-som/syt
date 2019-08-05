import sytlog
EN = 'en'
IN = 'in'
JP = 'jp'
ZD = 'zd'
HW = 'hw'
CN = 'cn'
TH = 'th'
VN = 'vn'

# _json_root = "E:/st_kf/doc/language"
# _excel_root = "E:/sy_translation"
_cfg = None
_cur = None
def setcfg(cfg):
    global _cfg
    _cfg = cfg

def choose(vsn):
    global _cur
    _cur = _cfg[vsn]
    sytlog.log(u"选择版本[%s]\n" % _cur['name'])

def json_root():
    if _cur == None:
        return "./"
    return _cur['json_root']

def excel_root():
    if _cur == None:
        return "./"
    return _cur['excel_root']