"""
●概要
端末のハードウェア・OS・インストールされているアプリケーション情報をWMIを用いて取得し、
ファイルサーバ上の指定されたフォルダにあるCSVファイルへ追記する。

実際にWMIから取得される情報は、CSV出力されるもの以外にも多数あるが、現在のところ捨てている。(2018/10/31)

出力先のCSVファイルはinventry.iniファイルに指定された情報を元にしているが、コマンドライン
パラメーターで指定することもできる。

コマンドラインパラメーターが指定された場合、iniファイルの情報よりも優先される。（2018/1031、未実装）

inventry.iniについては後述する。

●使い方
Usage: python inventry.py [option]
options:
--help
--version
--server
--drive
--path
--csv
--log LEVEL

●出力情報
CSV出力される情報は下記の通り。
・端末名
・メーカー名
・機種モデル
・OS・エディション
・OSのビルド番号（Windows10でのバージョンに相当）
・OSのビット数
・CPU名
・CPUモデル
・CPUクロック周波数
・登載メモリ（GB単位）
・登載ドライブ容量（GB単位、一つ目に認識されたドライブのみ）
・ネットワークカード名（一つ目）
・ネットワークカードのMACアドレス（一つ目）
・ネットワークカードに割り当てられたIPv4アドレス（一つ目）
・ネットワークカード名（二つ目。存在すれば）
・ネットワークカードのMACアドレス（二つ目。存在すれば）
・ネットワークカードに割り当てられたIPv4アドレス（二つ目。存在すれば）

●inventry.ini
ServerInfoセクションでは下記を設定する。
・DriveCaption  : ネットワークドライブのキャプション 
・ServerIP      : 共有フォルダを公開しているサーバのIPアドレス
・UserID        : 共有フォルダのアカウント。base64でエンコード済みの文字列を指定
・Passwd        : 共有フォルダのパスワード。base64でエンコード済みの文字列を指定
・ServerPath    : 共有フォルダのフルパス
・CSVFileName   : csvファイル名

Loggingセクションでは下記を設定する。
・LogLevel      : ログレベル。DEBUG, INFO, WARNING, ERROR, CRITICALのいずれかを指定する。

●注意点
ネットワークカードはスクリプトでは無線か有線かの区別はできず、WMIで列挙された順にCSVへ出力される。
そのためメンテナーはネットワークカード名、端末タイプ、無線LAN子機の存在等を考慮しつつ、
台帳へ反映させなければならない。
"""

import wmi, winreg
import configparser, argparse, logging
import base64, datetime, os, re

VERSION = "0.9.0"                   # バージョン
CONFIG_FILE_NAME = "inventry.ini"   # 設定ファイル名
CONFIG_SEG_SERVER = "ServerInfo"    # 設定ファイル内のサーバー情報セクション名
CONFIG_SEG_LOGGING = "Logging"      # 設定ファイル内のログセクション名

GIGABYTE = 1024 * 1024 * 1024       # ギガバイト定数

class InventryCfg():
    """
    INIファイルを読み込み、設定情報を保持する

    Attributes
    ----------
    _config : configparser
        iniファイルの情報を保持するconfigparserクラスのインスタンス
    """
    def __init__(self):
        """
        コンストラクタ
        """
        self._config = configparser.ConfigParser()
        self._config.sections()
        self._config.read(CONFIG_FILE_NAME)
        self._config[CONFIG_SEG_SERVER]['UserID'] = base64.standard_b64decode(self._config[CONFIG_SEG_SERVER]['UserID']).decode()
        self._config[CONFIG_SEG_SERVER]['Passwd'] = base64.standard_b64decode(self._config[CONFIG_SEG_SERVER]['Passwd']).decode()

    def cfg(self):
        """
        configparserクラスのインスタンスを返す

        Returns
        -------
        self._config : configparser
            iniファイル情報を保持するクラスのインスタンス
        """
        return self._config

# コマンドラインパラメータ設定・解析
def parseParam():
    parser = argparse.ArgumentParser()
    parser.add_argument('--version', action='version', default="", version='%(prog)s ' + VERSION)
    #parser.add_argument('--server', action='server', default="", server='')
    #parser.add_argument('--drive', action='drive', default="", drive='')
    #parser.add_argument('--path', action='path', default="", path='')
    #parser.add_argument('--csv', action='csv', default="", csv='')
    #parser.add_argument('--log', action='log', default="", log='')
    args = parser.parse_args()
    if args.version != '':
        print(args.version)
        exit()


# 設定ファイル（iniファイル）読み込み
def loadIni():
    config = configparser.ConfigParser()
    config.sections()
    try:
        config.read(getIniFilePath())
        config[CONFIG_SEG_SERVER]['UserID'] = base64.standard_b64decode(config[CONFIG_SEG_SERVER]['UserID']).decode()
        config[CONFIG_SEG_SERVER]['Passwd'] = base64.standard_b64decode(config[CONFIG_SEG_SERVER]['Passwd']).decode()
        return config
    except:
        print(getIniFilePath())
        print('iniファイルが読み込めませんでした。')
        exit(0)


def recordInfo(strCSV, csvFullPath):
    fp = open(csvFullPath, 'a')
    fp.write(strCSV)
    fp.close()
    getLogger().info('データ出力に成功しました。')


class WMIController():
    """
    WMIを用いて各種の必要な情報を取得する

    Attributes
    ----------
    _w : wmi
        wmiクラスのインスタンス
    """
    def __init__(self):
        """
        コンストラクタ
        """
        self._w = wmi.WMI()

    def operatingSystem(self):
        """
        OS情報を取得する

        Returns
        -------
        dictionary
            OS名、OSのビット数、OSのビルド番号
        """
        for os in self._w.Win32_OperatingSystem():
            return {"Caption": os.Caption.replace("Microsoft ", ""), "OSArchitecture": os.OSArchitecture.replace(" ビット", ""), "Version": os.Version}

    def computerSystem(self):
        """
        端末の物理的な情報を取得する

        Returns
        -------
        dictionary
            メーカー、端末モデル、ホスト名、所属ワークグループ、物理メモリ（GB単位）
        """
        items = self._w.Win32_ComputerSystem()
        for item in items:
            return {"Manufacturer": item.Manufacturer.strip(), "Model": item.Model.strip(), "DNSHostName": item.DNSHostName, "Workgroup": item.Workgroup, "TotalPhysicalMemory": int(item.TotalPhysicalMemory) / (1000*1000*1000)}

    def logicalDisk(self):
        """
        全ての論理ドライブの情報を取得する

        Returns
        -------
        array <- dictionary
            ドライブ名、ディスク総容量、ディスク空き容量
        """
        value = []
        items = self._w.Win32_LogicalDisk()
        for item in items:
            if item.DriveType == 3:
                value.append({"Caption": item.Caption.strip(), "Size": item.Size, "FreeSpace": item.FreeSpace})
        return value

    def diskDrive(self):
        """
        全ての物理ドライブの情報を取得する

        Returns
        -------
        array <- dictionary
            ディスク機種、シリアル番号、ディスク総容量(GB単位)
        """
        value = []
        items = self._w.Win32_DiskDrive()
        for item in items:
            value.append({"Model": item.Model.strip(), "SerialNumber": item.SerialNumber.strip(), "Size": int(item.Size) / GIGABYTE})
        return value

    def processor(self):
        """
        CPU情報を取得する

        Returns
        -------
        dictionary
            CPU機種、CPUモデル、クロック周波数
        """
        items = self._w.Win32_Processor()
        for item in items:
            data = item.Name.replace("(R)", "").replace("(TM)", "").replace("Intel ", "").replace("CPU", "").replace("-", " ")
            #print(data)
            m = re.match(r'(.+)\s+(.+)\s+@\s+(.+)', data)
            cpu = m.group(1).strip()
            model = m.group(2).strip()
            clock = m.group(3).strip()
            getLogger().debug("{:s}, {:s}, {:s}".format(cpu, model, clock))
            return {"CPU": cpu, "model": model, "clock": clock}

    def networkAdapterConfiguration(self):
        """
        ネットワークアダプタ情報を取得する

        Returns
        -------
        array <- dictionary
            アダプタ名、MACアドレス、IPアドレス（割り当てられていれば）

        Reference
            __getIpv4Addr
                IPv4とIPv6が格納されている配列から、IPv4のみを返す
        """
        value = []
        items = self._w.WIN32_NetworkAdapterConfiguration()
        for item in items:
            getLogger().debug(item.Caption)
            if re.match(r'.*Virtual.*', item.Caption):
                continue
            elif re.match(r'.*WAN\s+Miniport.*', item.Caption):
                continue
            elif re.match(r'.*Microsoft\sKernel\sDebug\s+.*', item.Caption):
                continue
            elif item.MACAddress is None:
                continue
            m = re.match(r'\[\d+\]\s+(.+)', item.Caption)
            caption = m.group(1)
            ip = self.__getIpv4Addr(item.IPAddress)
            value.append({"Caption": caption, "MACAddress": item.MACAddress, "IPAddress": ip})
        return value

    def __getIpv4Addr(self, addresses):
        """
        IPv4とIPv6が格納されている配列から、IPv4のみを返す

        Paramaters
        ----------
        addresses : array
            MACアドレスに関連付けられていたIPv4とIPv6のアドレスが格納されている配列

        Returns
        -------
            addr : string
                IPv4アドレス
        """
        getLogger().debug(addresses)
        if addresses is None:
            return ""

        for addr in addresses:
            if re.match(r'\d{1,3}\.\d{1,3}.\d{1,3}.\d{1,3}', addr):
                return addr
        
        return ""


class Machine():
    """
    端末の情報を取得・保持する

    Attributes
    ----------
    _w : WMIController
        WMIControllerクラスのインスタンス
    applications : array
        端末にインストールされている（ほぼ）全てのアプリケーションを保持
    machine : dictionary
        端末のハードウェア、OS関連情報を保持
    """
    def __init__(self):
        """
        コンストラクタ
        """
        self._w = WMIController()
        self.applications = []
        self.machine = {}

    def __getUninstallerEntry(self, root, path):
        """
        レジストリを走査し、インストールされているアプリケーション名を取得・保持する

        Paramaters
        ----------
        root : string
            レジストリのルート
        path : string
            走査するレジストリのパス

        Note
        ----
        WindowsUpdateによってインストールされているパッチ類やランタイムライブラリ、ビルドツール等は除外している。
        """
        index = 0
        try:
            key = winreg.OpenKey(root, path, 0, winreg.KEY_READ)
            while 1:
                name = winreg.EnumKey(key, index)
                sub = winreg.OpenKey(key, name, 0, winreg.KEY_READ)
                try:
                    value = winreg.QueryValueEx(sub, "DisplayName")
                except:
                    pass
                else:
                    if re.match(r'.*Update.*', value[0]):
                        pass
                    elif re.match(r'\(KB\d+\)', value[0]):
                        pass
                    elif re.match(r'.*Service Pack.*', value[0]):
                        pass
                    elif re.match(r'.*SDK.*', value[0]):
                        pass
                    elif re.match(r'.*Visual C\+\+', value[0]):
                        pass
                    elif re.match(r'WinRT.*', value[0]):
                        pass
                    elif re.match(r'.*Build Tools.*', value[0]):
                        pass
                    elif re.match(r'.*Development Kit.*', value[0]):
                        pass
                    elif len(value[0]) == 0:
                        pass
                    else:
                        self.applications.append(value[0])
                index += 1
        except OSError:
            pass


    def get(self):
        """
        端末のハードウェア・ソフトウェア情報を取得する
        """
        self.machine['OS'] = self._w.operatingSystem()
        self.machine['System'] = self._w.computerSystem()
        self.machine['LogicalDisk'] = self._w.logicalDisk()
        self.machine['DiskDrive'] = self._w.diskDrive()
        self.machine['Processor'] = self._w.processor()
        self.machine['Network'] = self._w.networkAdapterConfiguration()
        getLogger().debug(self.machine)

        self.__getUninstallerEntry(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
        self.__getUninstallerEntry(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
        self.__getUninstallerEntry(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
        getLogger().debug(self.applications)

    def outputCSV(self):
        """
        保持しているハードウェア・OSの情報をcsvフォーマットで返す

        Returns
        -------
        csv : string
            csv形式の端末のハードウェア・OS情報
        """
        csv = '{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}'.format(
            self.machine['System']['DNSHostName'],          # 端末名
            self.machine['System']['Manufacturer'],         # メーカー
            self.machine['System']['Model'],                # 機種
            self.machine['OS']['Caption'],                  # OS名
            self.machine['OS']['Version'],                  # OSバージョン（ビルド番号）
            self.machine['OS']['OSArchitecture'],           # OSビット数
            self.machine['Processor']['CPU'],               # CPU名
            self.machine['Processor']['model'],             # CPUモデル
            self.machine['Processor']['clock'],             # クロック周波数
            round(self.machine['System']['TotalPhysicalMemory']),   # 物理メモリ
            round(self.machine['DiskDrive'][0]['Size']),    # 物理ドライブ総容量
            self.machine['Network'][0]['Caption'],          # ネットワークカード名
            self.machine['Network'][0]['MACAddress'],       # MACアドレス
            self.machine['Network'][0]['IPAddress'],        # IPアドレス
            self.machine['Network'][1]['Caption'] if len(self.machine['Network']) >= 2 else "",     # ネットワークカード名（2枚目）
            self.machine['Network'][1]['MACAddress'] if len(self.machine['Network']) >= 2 else "",  # MACアドレス（2枚目）
            self.machine['Network'][1]['IPAddress'] if len(self.machine['Network']) >= 2 else "",   # IPアドレズ（2枚目）
        )
        return csv


def getLogger():
    return logging.getLogger(__name__)


def getLogFilePath():
    #dirname = os.path.dirname(__file__)
    #fname = re.search(r'(.+)\.py$', os.path.basename(__file__))
    #return (dirname + "\\" if len(dirname) else "") + fname.group(1) + ".log"
    return "C:\\Users\\admin\\inventry.log"

def getIniFilePath():
    #dirname = os.path.dirname(__file__)
    #return (dirname + "\\" if len(dirname) else "") + CONFIG_FILE_NAME
    return "C:\\Users\\admin\\" + CONFIG_FILE_NAME

def main():
    parseParam()        # コマンドラインパラメーター解析
    config = loadIni()  # iniファイル読み込み

    # ログファイル設定
    logging.basicConfig(filename=getLogFilePath(), format='[%(asctime)s] %(message)s', level=config[CONFIG_SEG_LOGGING]['LogLevel'])

    try:
        # 端末情報を取得
        device = Machine()
        device.get()
        csvStr = datetime.date.today().isoformat() + ", " + device.outputCSV() + "\n"
        getLogger().debug(csvStr)

        # あらかじめネットワークドライブを解除しておく
        # コマンドライン上の出力は表示させない
        cmd = 'net use ' + config[CONFIG_SEG_SERVER]['DriveCaption'] + ': /d > nul 2>&1'
        os.system(cmd)

        # csv出力先の共有フォルダをネットワークドライブとして接続
        cmd = 'net use ' + config[CONFIG_SEG_SERVER]['DriveCaption'] + ': \\\\' + config[CONFIG_SEG_SERVER]['ServerIP'] + config[CONFIG_SEG_SERVER]['ServerPath'] + ' /user:' + config[CONFIG_SEG_SERVER]['UserID'] + ' ' + config[CONFIG_SEG_SERVER]['Passwd']
        os.system(cmd)

        # csvファイルに出力
        recordInfo(csvStr, config[CONFIG_SEG_SERVER]['DriveCaption'] + ":\\" + config[CONFIG_SEG_SERVER]['CSVFileName'])

    finally:
        # ネットワークドライブ解除
        cmd = 'net use ' + config[CONFIG_SEG_SERVER]['DriveCaption'] + ': /d'
        os.system(cmd)
        logging.shutdown()


if __name__ == '__main__':
    main()
