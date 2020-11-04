from os.path import join

from Application import ReceivingOrders, AssetRegistration

ExcelPath = input('请将准备好的Excel文件拖动到此窗口中：')

# # 实例化接单模块
runRO = ReceivingOrders
runRO.ReceivingOrders()


# 分配资产实例化
runAR = AssetRegistration
runAR.AssetRegistraton(ExcelPath)