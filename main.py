# !/usr/bin/env python
# -*-coding:utf-8 -*-
import time
import traceback
import os
from threading import Thread
import re

import requests
from DrissionPage import ChromiumPage, ChromiumOptions
from loguru import logger
from pathlib import Path
from queue import Queue

proceed_node = set()
proceed_files = set()

req_queue = Queue()
download_queue = Queue()

# 配置项
# 公司ID
corpId = ""
# 组织（库）ID
target_orgid = ""
# 无权限文件
no_right_files = []
loggined_done = False



def process_download():
    while True:
        res = download_queue.get(block=True)
        if not res:
            continue
        try:
            node_info, url, headers, cookies, save_path, save_name = res
            p = Path(save_path)
            # 创建文件夹
            os.makedirs(p.absolute(), exist_ok=True)
            cookies = {x["name"]: x["value"] for x in cookies}
            download_success = False
            for retry_times in range(10):
                try:
                    req = requests.request(method="get", url=url, headers=headers, cookies=cookies)
                    filename = str(url).split("?")[0].split("/")[-1]
                    save_path = p.joinpath(filename)
                    if req.status_code == 200:
                        with open(save_path.absolute(), "wb") as f:
                            f.write(req.content)
                        download_success = True
                        break
                    else:
                        raise Exception(f"下载失败，返回状态码{req.status_code},内容：{req.content}")
                except Exception as e:
                    logger.error(f"下载文件{url} 重试{retry_times} 出错：{e}")
                    time.sleep(5)
                    continue
            if not download_success:
                logger.error(f"下载文件{url}失败，推回节点到浏览器进行重试")
                q.put(node_info)
        except Exception as e:
            logger.error(f"下载{res}出错 {e}：{traceback.format_exc()}")

def request_repeater(q):
    while True:
        res = req_queue.get(block=True)
        data = None
        if res:
            if "/dentry/list?" not in str(res.url):
                logger.info(f"跳过{str(res.url)}")
                continue
            for _ in range(10):
                try:
                    logger.info(f"二次请求{res.url}，待请求长度：{req_queue.qsize()}")
                    cookies = {x["name"]: x["value"] for x in res.request.cookies}
                    data = requests.request(method="get", url=res.url, headers=res.request.headers,
                                            cookies=cookies)
                    data = data.json()["data"]
                    logger.info(f"二次请求完成，待请求长度：{req_queue.qsize()}")
                    break
                except Exception as e:
                    logger.error(f"二次请求{res.url} 出错：{e}")
                    time.sleep(5)
                    continue
            if data:
                process_req(q, data)
            else:
                logger.error(f"二次请求{res.url} 失败次数超过10，放弃")


def process_req(q, data):
    if not data:
        return
    if "children" in data:
        process_node_name = data['name']
        item_list = data["children"]
        added_names = []
        for node_info in item_list:
            node_name = node_info['name']
            node_uuid = node_info['dentryUuid']
            if node_uuid not in proceed_node:
                added_names.append(node_name)
                q.put(node_info)
                proceed_node.add(node_uuid)
        if added_names:
            logger.info(f"队列长度：{q.qsize()} 从【{process_node_name}】 添加子节点{len(added_names)}个：{', '.join(added_names)}")

class Processer:

    def __init__(self, q, index=0):
        self.idx = index
        self.q = q
        self.page = ChromiumPage(ChromiumOptions().set_local_port(int(f"933{index}")).set_user_data_path(f'data{index}'))
        package_urls = ['box/api/v2/dentry/list?']
        self.page.listen.start(package_urls, res_type=True)
        self.page.get(f'https://alidocs.dingtalk.com/i/desktop/spaces/?corpId={corpId}')
        self.inited = False
        self.headers = {}
        self.cookies = {}

    def run(self):
        empty_count = 0
        while True:
            if loggined_done and not self.inited:
                self.inited = True
                # 打开组织页面
                self.page.get(f'https://alidocs.dingtalk.com/i/spaces/{target_orgid}/overview?corpId={corpId}')
            self.block_wait()
            while self.page.listen._caught.qsize():
                res = self.page.listen.wait(timeout=5)
                if res:
                    if res.response.body and res.response.body.get("data"):
                        data = res.response.body["data"]
                        process_req(self.q, data)
                        try:
                            # 更新最新header以及cookies
                            self.headers = res.request.headers
                            self.cookies = res.request.cookies
                        except Exception:
                            pass

                    else:
                        req_queue.put(res)

            if not self.q.empty():
                item = self.q.get()
                for retry in range(4):
                    try:
                        self.process_node(item)
                        break
                    except Exception as e:
                        logger.error(f"处理{item}时发生错误：{e} 重试{retry+1}")
                empty_count = 0
                continue

            if empty_count > 30:
                logger.info(f"[{self.idx}] 退出")
                break
            empty_count += 1
            time.sleep(5)

        self.page.close()
        self.page.browser.quit()

    def block_wait(self):
        time.sleep(1)
        while not self.page.listen.wait_silent(targets_only=True):
            time.sleep(1)

    def process_node(self, node_info, load_page=True):
        node_name = node_info['name']
        node_uuid = node_info['dentryUuid']
        ancestorList = node_info['ancestorList']
        # if node_uuid in proceed_node:
        #     logger.info(f"[{self.idx}]jump{node_uuid}")
        #     return

        parent_node_name = "根节点"
        if ancestorList:
            parent_node_name = ancestorList[-1]['name']
        logger.info(f"[{self.idx}] 开始处理节点:{node_name} 父节点：{parent_node_name}")
        # 直接跳转页面
        if load_page:
            self.page.get(f"https://alidocs.dingtalk.com/i/nodes/{node_uuid}")
        self.block_wait()
        # 判断是否页面白屏
        if node_info.get('contentType') == 'alidoc' or node_info.get('dentryType') == 'file':
            logger.info(f"[{self.idx}] {node_name}是文件，继续处理")
            success = self.process_file(node_info)
            if not success:
                logger.info(f"[{self.idx}] {node_name} 文件 处理失败，推回队列 后续重试")
                self.q.put(node_info)
            # 选中节点
            find_div = f"@data-rbd-draggable-id={node_uuid}"
            try:
                item = self.scroll_to_see(find_div)
                self.to_item(item)
            except Exception as e:
                logger.info(f"[{self.idx}] {find_div}: {e} {traceback.format_exc()}")
                self.process_node(node_info, load_page=False)

        else:
            # 选中节点
            find_div = f"@data-rbd-draggable-id={node_uuid}"
            try:
                button = self.scroll_to_see(find_div)
                if not button:
                    self.process_node(node_info)
                time.sleep(0.5)
                self.to_item(button)
                button.click()
                time.sleep(0.5)
            except Exception as e:
                logger.info(f"[{self.idx}] {find_div}: {e} {traceback.format_exc()}")
                self.process_node(node_info, load_page=False)
    def to_item(self, item):
        item_loc = item.rect.location
        self.page.scroll.to_location(item_loc[0], item_loc[1])

    def check_alert(self):
        for i in range(2):
            # 检查是否有按钮 继续导出
            has_limit = self.page.eles("tag:button@@text():继续导出", timeout=0.5)
            if has_limit:
                has_limit[0].click()
                break

    def scroll_to_see(self, loc,retry_times=0):
        if retry_times > 5:
            return
        try:
            # 先尝试直接找
            client_now = self.page.run_js(
                'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollTop')
            client_height = self.page.run_js(
                'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].clientHeight')
            tree = self.page.ele(".:MAINSITE_CATALOG-node-tree-list")
            item = self.scroll(tree, client_now, client_height, loc)
            if item:
                return item
            else:
                return self.scroll_to_see(loc,retry_times+1)
        except Exception as e:
            logger.error(f"滚动时出错：{e}，等待重试")
            return self.scroll_to_see(loc,retry_times+1)

    def scroll(self, tree, start, client_height, loc):
        if start == 0:
            tree.scroll.to_top()
            tree.scroll.to_location(0, 0)
            time.sleep(0.5)
        item = self.page.ele(loc, timeout=2)
        if item:
            return item
        last_scrollTop = None
        start_height = int(start)
        to_height = self.page.run_js(
            'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollHeight')
        # 每次滚动的高度
        roll_height = 300
        while start_height < to_height:
            to_height = self.page.run_js(
                'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollHeight')
            if last_scrollTop is not None and last_scrollTop == to_height and (
                    start_height > (to_height / 2)):
                return self.scroll(tree, 0, client_height, loc)
            last_scrollTop = self.page.run_js(
                'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollTop')
            logger.info(f"[{self.idx}] 正在滚动【高度信息：当前：{last_scrollTop or 0}->目标：{start_height}/整体{to_height}】")
            start_height += roll_height
            tree.scroll.to_location(0, start_height)
            item = self.page.ele(loc, timeout=3)
            if item:
                return item

    def process_file(self, node_info, retry_times=0):
        def c(x):
            x = (x or "").replace('\\', '_').replace(' ', '_').replace(':', '_')
            x = x.replace('/', '_').replace('?', '_').replace("*", "_")
            x = x.replace('\n', '_').strip()
            x = re.sub(r"(?u)[^-\w.]", "", x)
            return x
        node_name = node_info['name']
        file_type = node_name.split(".")[-1]
        node_uuid = node_info['dentryUuid']
        ancestor_path = [c(x['name']) for x in node_info['ancestorList']]
        if node_uuid in proceed_files:
            return True
        proceed_files.add(node_uuid)
        self.block_wait()
        file_path = '\\'.join([c(x) for x in ancestor_path])
        node_name = c(node_name.rsplit(".", 1)[0])
        logger.info(f"[{self.idx}] 处理文件:{node_name} 路径：{file_path} 文件类型：{file_type}")
        path = Path(".").absolute()
        path = path.joinpath(target_orgid)
        path = path.joinpath(file_path)
        os.makedirs(str(path.absolute()), exist_ok=True)
        fname = path.joinpath(node_name)
        if fname.exists() and fname.is_dir() and len(list(os.listdir(fname))) > 0:
            logger.info(f"[{self.idx}] 节点已完成下载：{fname} 跳过。")
            return True
        if retry_times > 2:
            self.page.refresh()
        if retry_times > 5:
            logger.error(f"请求节点：{node_name}，{ancestor_path}出错次数超过10次，放弃")
            no_right_files.append((file_path, node_name, file_type))
            proceed_files.remove(node_uuid)
            return
        # 选中节点
        find_div = f"@data-rbd-draggable-id={node_uuid}"
        try:
            item = self.scroll_to_see(find_div)
            self.to_item(item)
            item.click()
        except Exception as e:
            logger.info(f"[{self.idx}] {find_div}: {e}")
            return self.process_file(node_info, retry_times+1)
        # 判断是否无权限访问
        notice_eles = self.page.eles("@data-item-key=apply-title-view") or []
        for ne in notice_eles:
            if "暂无权限访问" in str(ne.text):
                no_right_files.append((file_path, node_name, file_type))
                logger.info(f"[{self.idx}] 节点：{node_name} 无访问权限，跳过")
                return True

        # 如果是链接
        if file_type == "dlink":
            file_type = node_info['linkSourceInfo']['extension']
            logger.info(f"[{self.idx}] 链接文件：{fname} 真实文件类型为：{file_type}")
        self.page.set.download_path(str(fname.absolute()))
        self.page.set.download_file_name(node_name)
        self.page.set.when_download_file_exists("skip")
        time.sleep(5)
        try:
            download_task = False
            last_err = None
            if "adoc" in file_type:
                limited_toolbar = self.page.eles("@data-testid=doc-header-more-button", timeout=2)
                if limited_toolbar:
                    for i in range(5):
                        try:
                            limited_toolbar[0].click()
                            time.sleep(0.5)
                            self.page.ele("@data-item-key=export").click()
                            self.page.ele("@data-item-key=exportAsWord").click()
                            self.check_alert()
                            download_task = self.page.wait.download_begin(timeout=120)
                            last_err = None
                            break
                        except Exception as err: 
                            last_err = err
                            time.sleep(3)
                            continue
                else:
                    for i in range(5):
                        try:
                            normal_toolbar = self.page.eles("@data-testid=bi-toolbar-menu", timeout=2)
                            if normal_toolbar:
                                normal_toolbar[0].click()
                                self.page.ele("@data-testid=bi-toolbar-menu").click()
                                time.sleep(0.5)
                                self.page.ele("@data-testid=menu-item-J_file").click()
                                time.sleep(0.5)
                                self.page.ele("@data-testid=menu-item-J_fileExport").click()
                                time.sleep(0.5)
                                self.page.ele("@data-testid=menu-item-J_exportAsWord").ele("text:Word").click()
                                self.check_alert()
                                download_task = self.page.wait.download_begin(timeout=120)
                                last_err = None
                                break
                        except Exception as err: 
                            last_err = err
                            time.sleep(3)
                            continue
                    else:
                        no_right_files.append((file_path, node_name, file_type))
            elif "axls" in file_type:
                limited_toolbar = self.page.eles("@data-testid=doc-header-more-button", timeout=2)
                if limited_toolbar:
                    for i in range(5):
                        try:
                            limited_toolbar[0].click()
                            time.sleep(0.5)
                            self.page.ele("@data-item-key=DOWNLOAD_AS").click()
                            self.page.ele("@data-item-key=EXCEL").click()
                            self.check_alert()
                            download_task = self.page.wait.download_begin(timeout=120)
                            last_err = None
                            break
                        except Exception as err: 
                            last_err = err
                            time.sleep(3)
                            continue
                else:
                    for i in range(5):
                        try:
                            normal_toolbar = self.page.eles("#wiki-new-sheet-iframe")
                            if normal_toolbar:
                                normal_toolbar[0].ele(
                                    "@data-testid=submenu-menubar-table").ele("text:表格").click()
                                time.sleep(0.5)
                                self.page.ele("#wiki-new-sheet-iframe").ele(
                                    "@data-testid=submenu-export-excel").ele("text:下载为").click()
                                time.sleep(0.5)
                                self.page.ele("#wiki-new-sheet-iframe").ele("text:Excel").click()
                                self.check_alert()
                                download_task = self.page.wait.download_begin(timeout=120)
                                last_err = None
                                break
                        except Exception as err: 
                            last_err = err
                            time.sleep(3)
                            continue
                    else:
                        for i in range(5):
                            try:
                                download_button = self.page.eles("@data-item-key=download", timeout=1)
                                if download_button:
                                    download_button[0].click()
                                    self.check_alert()
                                    download_task = self.page.wait.download_begin(timeout=120)
                                    last_err = None
                                    break
                                else:
                                    no_right_files.append((file_path, node_name, file_type))
                            except Exception as err: 
                                last_err = err
                                time.sleep(3)
                                continue

            else:
                logger.error(f"[{self.idx}] 跳过未处理的文件下载：{fname} ")
                return True
            if last_err:
                raise last_err
            need_restart = False
            if not download_task:
                need_restart = True
            else:
                # 等待下载
                while not download_task.is_done:
                    time.sleep(.5)
                if not download_task.final_path and not download_task.state == "skipped":
                    logger.info(f"[{self.idx}] 下载{fname} 任务失败 任务最终状态：{download_task.state}")
                    if "blob" not in download_task.url:
                        res = (node_info, download_task.url, self.headers, self.cookies, str(fname.absolute()), node_name)
                        download_queue.put(res)
                        logger.info(f"[{self.idx}] 生成下载{fname}任务")
            if need_restart:
                logger.error(f"[{self.idx}] 下载：{fname} 未完成任务生成就结束了，重试一次")
                proceed_files.remove(node_uuid)
                return self.process_file(node_info, retry_times + 1)
        except Exception as e:
            logger.error(f"[{self.idx}] 下载：{fname} 时出现问题，可能是无下载权限造成的：{e} {traceback.format_exc()}")
            # no_right_files.append((file_path, node_name, file_type))
            proceed_files.remove(node_uuid)
            return self.process_file(node_info, retry_times+1)
        logger.info(f"[{self.idx}] 已完成处理文件:{node_name} 路径：{file_path} 文件类型：{file_type}，等待..")
        time.sleep(0.5)
        return True


if __name__ == "__main__":
    threads = []
    q = Queue()
    logger.info("启动浏览器。。。")
    for i in range(5):
        thread = Thread(target=request_repeater, args=(q,))
        thread.start()

    for i in range(5):
        thread = Thread(target=process_download, args=())
        thread.start()

    for i in range(5):
        thread = Thread(target=Processer(q, i).run, args=())
        thread.start()
    input(f"请完成所有浏览器的登录，并在完成后任意键继续")
    loggined_done = True
    input(f"全部下载完成后任意键继续")


    [x.join() for x in threads]
    time.sleep(5)
    while q.qsize():
        time.sleep(10)
    print(f"无权限文件的列表\n{no_right_files}")
    input("全部抓取完成，任意键退出")

