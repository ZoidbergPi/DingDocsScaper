# !/usr/bin/env python
# -*-coding:utf-8 -*-
import time
import traceback
import os
from DrissionPage import ChromiumPage
from loguru import logger
from pathlib import Path

page = ChromiumPage()

# 配置项
# 公司ID
corpId = ""
# 组织（库）ID
target_orgid = ""
# 无权限文件
no_right_files = []

def c(x):
    return (x or "").replace('\\', '_').replace(' ', '_').replace('/', '_')


def process_file(node_info):
    node_name = node_info['name']
    file_type = node_name.split(".")[-1]
    node_uuid = node_info['dentryUuid']
    ancestor_path = [x['name'] for x in node_info['ancestorList']]
    if node_uuid in proceed_files:
        return
    proceed_files.add(node_uuid)
    file_path = '\\'.join([c(x) for x in ancestor_path])
    node_name = c(node_name.rsplit(".", 1)[0])
    logger.info(f"处理文件:{node_name} 路径：{file_path} 文件类型：{file_type}")
    path = Path(".").absolute()
    path = path.joinpath(target_orgid)
    path = path.joinpath(file_path)
    os.makedirs(str(path.absolute()), exist_ok=True)
    fname = path.joinpath(node_name)
    if fname.exists():
        logger.info(f"文件：{fname} 已存在 跳过。")
        return
    # 选中节点
    find_div = f"@data-rbd-draggable-id={node_uuid}"
    try:
        item = scroll_to_see(find_div)
        item.scroll.to_center()
        item.click()
    except Exception as e:
        logger.info(f"{find_div}: {e}")
    # 如果是链接
    if file_type == "dlink":
        file_type = node_info['linkSourceInfo']['extension']
        logger.info(f"链接文件：{fname} 真实文件类型为：{file_type}")
    # # 对页面进行截图留档
    # page.get_screenshot(folder_path, node_name + "_screenshot.jpg", full_page=True)
    page.set.download_path(str(fname.absolute()))
    page.set.download_file_name(node_name)
    try:
        if "adoc" in file_type:
            limited_toolbar = page.eles("@data-testid=doc-header-more-button", timeout=2)
            if limited_toolbar:
                limited_toolbar[0].click()
                time.sleep(0.5)
                page.ele("@data-item-key=export").click()
                page.ele("@data-item-key=exportAsWord").click()
                page.wait.download_begin()
            else:
                normal_toolbar = page.eles("@data-testid=bi-toolbar-menu", timeout=2)
                if normal_toolbar:
                    normal_toolbar[0].click()
                    page.ele("@data-testid=bi-toolbar-menu").click()
                    time.sleep(0.5)
                    page.ele("@data-testid=menu-item-J_file").click()
                    time.sleep(0.5)
                    page.ele("@data-testid=menu-item-J_fileExport").click()
                    time.sleep(0.5)
                    page.ele(
                        "@data-testid=menu-item-J_exportAsWord").ele("text:Word").click()
                    page.wait.download_begin()
                else:
                    no_right_files.append((file_path, node_name, file_type))
        elif "axls" in file_type:
            limited_toolbar = page.eles("@data-testid=doc-header-more-button", timeout=2)
            if limited_toolbar:
                limited_toolbar[0].click()
                time.sleep(0.5)
                page.ele("@data-item-key=DOWNLOAD_AS").click()
                page.ele("@data-item-key=EXCEL").click()
            else:
                normal_toolbar = page.eles("#wiki-new-sheet-iframe")
                if normal_toolbar:
                    normal_toolbar[0].ele(
                        "@data-testid=submenu-menubar-table").ele("text:表格").click()
                    time.sleep(0.5)
                    page.ele("#wiki-new-sheet-iframe").ele(
                        "@data-testid=submenu-export-excel").ele("text:下载为").click()
                    time.sleep(0.5)
                    page.ele("#wiki-new-sheet-iframe").ele("text:Excel").click()
                    page.wait.download_begin()
                else:
                    no_right_files.append((file_path, node_name, file_type))
    except Exception as e:
        logger.error(
            f"下载：{fname} 时出现问题，可能是无下载权限造成的：{e} {traceback.format_exc()}")
    logger.info(f"已完成处理文件:{node_name} 路径：{file_path} 文件类型：{file_type}，等待..")
    time.sleep(0.5)


def scroll(tree, start, to, client_height, loc):
    if start == 0:
        tree.scroll.to_top()
        tree.scroll.to_location(0, 0)
        time.sleep(0.5)
    item = page.ele(loc, timeout=2)
    if item:
        return item
    last_scrollTop = None
    for r in range(int(start), int(to), int(300)):
        logger.info(f"正在滚动，当前：{last_scrollTop or 0}->目标：{r}/整体{to}")
        scrollTop = page.run_js(
            'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollTop')
        if last_scrollTop is not None and last_scrollTop == scrollTop and (
                r > (to / 2)):
            return scroll(tree, 0, to, client_height, loc)
        last_scrollTop = int(scrollTop)
        if not r:
            r = int(to/2)
        tree.scroll.to_location(0, r)
        item = page.ele(loc, timeout=3)
        if item:
            return item


def scroll_to_see(loc):
    # 先尝试直接找
    height = page.run_js(
        'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollHeight')
    client_now = page.run_js(
        'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].scrollTop')
    client_height = page.run_js(
        'return document.getElementsByClassName("MAINSITE_CATALOG-node-tree-list")[0].clientHeight')
    tree = page.ele(".:MAINSITE_CATALOG-node-tree-list")
    item = scroll(tree, client_now, height, client_height, loc)
    if item:
        item.scroll.to_center()
        return item
    else:
        return scroll_to_see(loc)


def process_node(node_info, load_page=True):
    node_name = node_info['name']
    node_uuid = node_info['dentryUuid']
    ancestorList = node_info['ancestorList']
    if node_uuid in proceed_node:
        return
    proceed_node.add(node_uuid)
    parent_node_name = "根节点"
    if ancestorList:
        parent_node_name = ancestorList[-1]['name']
    logger.info(f"开始处理节点:{node_name} 父节点：{parent_node_name}")
    # 直接跳转页面
    if load_page:
        page.get(f"https://alidocs.dingtalk.com/i/nodes/{node_uuid}")
        time.sleep(2)
    if node_info.get('contentType') == 'alidoc' or node_info.get('dentryType') == 'file':
        logger.info(f"{node_name}是文件，继续处理")
        process_file(node_info)
        # 如果有展开按钮，就点击展开
        # 选中节点
        find_div = f"@data-rbd-draggable-id={node_uuid}"
        try:
            item = scroll_to_see(find_div)
            item.scroll.to_center()
            time.sleep(0.5)
            click_areas = item.eles("@data-testid=expand-click-area")
            if click_areas:
                click_areas[0].click()
                time.sleep(2)
        except Exception as e:
            logger.info(f"{find_div}: {e} {traceback.format_exc()}")
            process_node(node_info, load_page=False)

    else:
        # 选中节点
        find_div = f"@data-rbd-draggable-id={node_uuid}"
        try:
            button = scroll_to_see(find_div)
            if not button:
                process_node(node_info)
            time.sleep(0.5)
            button.scroll.to_center()
            button.click()
            time.sleep(0.5)
        except Exception as e:
            logger.info(f"{find_div}: {e} {traceback.format_exc()}")
            process_node(node_info, load_page=False)


if __name__ == "__main__":
    page.get(f'https://alidocs.dingtalk.com/i/desktop/spaces/?corpId={corpId}')
    page_url_now = page.url
    while page.url.startswith('https://login.dingtalk.com/oauth2/challenge.htm'):
        input(f"等待登录，完成后任意键继续")
    package_urls = ['box/api/v2/dentry/list']
    # 监听列表
    page.listen.start(package_urls)

    # 打开组织页面
    page.get(f'https://alidocs.dingtalk.com/i/spaces/{target_orgid}/overview?corpId={corpId}')
    dentry_mapping = {}
    proceed_node = set()
    proceed_files = set()

    while True:
        res = page.listen.wait(timeout=2)
        if res and res.response.body:
            data = res.response.body["data"]
            if "children" in data:
                item_list = data["children"]
                for item in item_list:
                    process_node(item)
        if not res:
            break
    print(f"无权限文件的列表\n{no_right_files}")
