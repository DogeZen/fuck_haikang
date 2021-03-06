'''
Win32_NetworkAdapterConfiguration类: https://docs.microsoft.com/zh-cn/windows/desktop/CIMWin32Prov/win32-networkadapterconfiguration#properties
'''
import wmi,ctypes,sys
objNicConfig = None
arrIPAddresses = ['192.168.1.100', '192.168.1.150']  # IP地址列表
arrSubnetMasks = ['255.255.255.0', '255.255.255.0']  # 子网掩码列表
arrDefaultGateways = ['192.168.1.253']  # 默认网关列表
arrGatewayCostMetrics = [1]  # 默认网关跳跃点
arrDNSServers = ['11.114.114.114', '8.8.8.8']  # DNS服务器列表
intReboot = 0


def GetNicConfig():
    wmiService = wmi.WMI()
    colNicConfigs = wmiService.Win32_NetworkAdapterConfiguration(IPEnabled=True)
    # 根据需要过滤出指定网卡，本只有一个网卡，不做过滤直接获取第一个
    index=0
    for obj in colNicConfigs:
        print(obj.Index)
        print(obj.Description)
        print(obj.SettingID)
        if(('Ethernet' in obj.Description or 'Connection' in obj.Description or 'PCIe' in obj.Description)):
            break
        index+=1
    if len(colNicConfigs) < 1:
        print('没有找到可用的网络适配器')
        return False
    else:
        global objNicConfig
        objNicConfig = colNicConfigs[index]
        return True


def SetAutoDNS():
    returnValue = objNicConfig.SetDNSServerSearchOrder()
    if returnValue[0] == 0:
        print('设置自动获取DNS成功')
    elif returnValue[0] == 1:
        print('设置自动获取DNS成功')
        intReboot += 1
    else:
        print('ERROR: DNS设置发生错误')
        return False
    return True


def SetAutoIP():
    returnValue = objNicConfig.EnableDHCP()
    if returnValue[0] == 0:
        print('设置自动获取IP成功')
    elif returnValue[0] == 1:
        print('设置自动获取IP成功')
        intReboot += 1
    else:
        print('ERROR: IP设置发生错误')
        return False
    return True


# 切换为静态IP
def EnableStatic():
    return SetIP() and SetGateways() and SetDNS()


# 切换为自动获取IP、DNS
def EnableDHCP():
    return SetAutoDNS() and SetAutoIP()


def main():
    if not GetNicConfig():
        return False


    else:
        print('正在切换为动态IP...')
        if EnableDHCP():
            if intReboot > 0:
                print('需要重新启动计算机')
            else:
                print('切换为动态DHCP成功!')
        else:
            print('请关闭控制面板、以管理员权限运行重试')


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if __name__ == "__main__":
    try:
        if is_admin():
            print("已使用管理员权限打开")
            main()
        else:
            print("未使用管理员权限打开，自动获取管理员权限")
            # Re-run the program with admin rights
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        c = input()
    except:
        print('未检测有设备接入有线网卡，请用网线连接摄像头和电脑')
        c = input()