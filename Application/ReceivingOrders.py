from selenium import webdriver
import os,time


# 接单模块
def ReceivingOrders(*args, **kwargs):
    if not os.path.exists('./File/list.txt'):
        print()
        print('=================================首次配置=================================')
        print()

        site = input("请输入办公点地址，如：“义乌福田银座大厦”，可填写“福田”：")
        while site == '':
            if site == '':
                site = input('办公点不允许为空，请重新输入：')
            else:
                with open(r'./File/list.txt', 'w+', encoding='utf8') as f:
                    f.write(str(site))
                    f.seek(0, 0)
                    f.close()
                    print()
                print('配置完成，地址是：', site)
                print()
                print('2秒后开始自动接单！！！')
                time.sleep(2)
                break
        with open(r'./File/list.txt', 'w+', encoding='utf8') as f:
            f.write(str(site))
            f.seek(0, 0)
            f.close()
        print('地址已保存，地址是：', site)
        print('2秒后开始自动接单！！！')
        time.sleep(2)
    else:
        with open('./File/list.txt', 'r+', encoding='utf8') as file:
            content = file.read()
            print()
            print('===================================================================================')
            print('||              读取默认配置地址，开始接单，请不要动浏览器！！！                 ||')
            print('||                                                                               ||')
            print('|| 如需重新配置地址，请手动删除‘ITAM/Flie/‘下的 ’list.txt’文件，再次执行脚本 ||')
            print('===================================================================================')
            print()
            print('读取配置文件地址是：', content)
            print()
            print('2秒后开始自动接单！！！')
            time.sleep(2)
            print()
            print('等待接单中..........')
            file.close()
        file.close()

    # 打开浏览器键入URL
    options = webdriver.ChromeOptions()

    try:
        options.add_argument(
            "user-data-dir=C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache")  # 携带cookice
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        wd = webdriver.Chrome(options=options, executable_path=r'./File/chromedriver.exe')
        wd.implicitly_wait(2000)  # 等待加载软件时间
        # wd.minimize_window()

        wd.get(
            'http://asset-pre.bytedance.net/real-asset/process/asset-appropriation?current=1&pageSize=200&subStatus=1001')
        wd.implicitly_wait(2000)  # 等待用户验证的时间

        # 抓取到数据然后点击区域按钮
        element = wd.find_elements_by_xpath(
            '//*[@id="app"]/section/section/main/div[2]/div/div[1]/div[1]/div[2]/div[6]')
        for i in element:
            i.click()

        # 进入到input按钮自动写入地址
        wd.implicitly_wait(500)  # 等待加载点击后出现地址的时间
        element = wd.find_elements_by_xpath('//*[@id="rc-tree-select-list_1"]/span/span/input')
        for i in element:
            with open('./File/list.txt', 'r', encoding='utf-8') as file:
                content = file.read()
                i.send_keys(content)

        # 点击搜索到的地址开始查询数据
        element = wd.find_elements_by_xpath('//*[@id="rc-tree-select-list_1"]/ul/li/ul/li/ul/li/ul/li')
        for i in element:
            i.click()

        # 点击工单来源按钮
        time.sleep(3)  # 等待加载点击后出现地址的时间
        element = wd.find_elements_by_xpath('//*[@id="app"]/section/section/main/div[2]/div/div[1]/div[1]/div[2]/div[5]/div/div/span/div/div')
        for i in element:
            i.click()

        time.sleep(1)  # 选择工单来源(入职领用)
        element = wd.find_elements_by_xpath(
            '/html/body/div[4]/div/div/div/div/ul/li[2]')
        for i in element:
            i.click()

        # 统计工单数量
        element = wd.find_element_by_xpath(
            '//*[@id="app"]/section/section/main/div[2]/div/div[2]/div/div/div/div/div/div[3]/div/div/table/tbody')
        with open(r'./File/element', 'w+', encoding='utf8') as f:
            f.write(str(element.text))
        num_lines = sum(1 for line in open(r'./File/element'))

        # 接单功能
        wd.implicitly_wait(13)  # 等待加载工单时间
        count = 0
        try:
            # 根据统计的工单数量进行接单
            while count <= num_lines:
                element = wd.find_elements_by_xpath("//button")
                wd.implicitly_wait(20)
                for element in element:
                    element.text
                try:
                    element.click()
                    count += 1
                    wd.implicitly_wait(10)
                    print('您一共有', num_lines, '单')
                    print('成功接单：', count, '单')
                    if count == num_lines:
                        print()
                        print('2秒后开始分配资产.......勿动~！')
                        wd.implicitly_wait(2)
                        wd.quit()
                        break
                except Exception as f:
                    if count == num_lines:
                        print()
                        print('############# 没有事情可做....... #############')
                        print()
                        print('进入自动分配资产模式......')
                        print()
                        wd.quit()
        except Exception as f:
            wd.quit()
    except Exception as f:
        print()
        print('############# 请关闭此脚本打开的 “谷歌浏览器” 后再执行此操作！！！#############')
        print()
        time.sleep(1)
        print('5秒后自动退出程序........')
        time.sleep(5)
        wd.quit()


