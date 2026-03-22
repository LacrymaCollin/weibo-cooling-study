import pandas as pd
import time
import os
import requests
import re
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# --------------------------------------------
SAVE_EXCEL = "weibo_collect_result_CI_社区中心_可加可不加.xlsx"
IMG_SAVE_DIR = "weibo_keyword_images_CI_社区中心_可加可不加"
WAIT_TIME = 5
KEYWORD_EXCEL_PATH = "keywords_CI_社区中心_可加可不加.xlsx"
COLLECT_PAGES_PER_KEYWORD = 10
MAX_EXCEL_ROWS = 1048500
# ---------------------------------------------------------------------

# 提前创建图片文件夹
if not os.path.exists(IMG_SAVE_DIR):
    os.makedirs(IMG_SAVE_DIR)


def init_driver():
    """初始化浏览器，深度模拟真实用户"""
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    # 模拟真实浏览器UA
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0")
    options.add_argument("--disable-blink-features=AutomationControlled")

    # 适配驱动路径
    try:
        service = Service()  # 自动查找系统驱动
    except:
        service = Service("./msedgedriver.exe")  # 同目录驱动

    driver = webdriver.Edge(service=service, options=options)
    # 隐藏自动化特征
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3]});
            Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN', 'zh']});
            window.navigator.chrome = {runtime: {}};
        """
    })
    return driver


def close_popups(driver):
    """关闭各类弹窗"""
    try:
        popups = [
            "//button[contains(text(), '关闭')]",
            "//button[contains(text(), '取消')]",
            "//span[contains(@class, 'close')]",
            "//div[contains(@class, 'mask')]//button"
        ]
        for popup_xpath in popups:
            try:
                popup_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, popup_xpath))
                )
                driver.execute_script("arguments[0].click();", popup_btn)
                time.sleep(0.5)
            except:
                continue
    except:
        pass


def read_keywords_from_excel(file_path):
    """从Excel读取关键词"""
    try:
        df = pd.read_excel(file_path)
        if "关键词" not in df.columns:
            raise ValueError("Excel文件必须包含「关键词」列！")
        keywords = df["关键词"].dropna().str.strip().unique().tolist()
        if not keywords:
            raise ValueError("Excel中无有效关键词！")
        print(f"✅ 成功读取 {len(keywords)} 个关键词：{keywords}")
        return keywords
    except Exception as e:
        print(f"❌ 读取关键词Excel失败：{e}")
        exit(1)


def direct_search_by_url(driver, keyword):
    """通过URL直接搜索（绕过搜索框交互）"""
    try:
        # 1. URL编码关键词（处理空格、中文）
        encoded_keyword = quote(keyword, encoding='utf-8')
        # 2. 拼接微博搜索URL
        search_url = f"https://s.weibo.com/weibo?q={encoded_keyword}&Refer=STopic_weibo"
        # 3. 访问搜索结果页
        driver.get(search_url)
        time.sleep(WAIT_TIME)

        # 4. 关闭弹窗
        close_popups(driver)

        # 5. 等待结果加载，检查是否无结果
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )

        # 检查无结果
        try:
            # 匹配微博无结果的多种提示
            no_result_elems = driver.find_elements(By.XPATH,
                                                   "//div[contains(text(), '没有找到相关结果') or contains(text(), '暂无结果') or contains(@class, 'noresult')]"
                                                   )
            if no_result_elems:
                print(f"⚠️ 关键词「{keyword}」无搜索结果，跳过")
                return False
        except:
            pass

        # 检查是否加载成功（是否有微博卡片）
        try:
            driver.find_element(By.CSS_SELECTOR, "div.card-wrap, div.feed-item")
            print(f"✅ 成功搜索关键词：{keyword}（URL方式）")
            return True
        except:
            print(f"⚠️ 关键词「{keyword}」搜索结果加载异常，跳过")
            return False

    except TimeoutException:
        print(f"⚠️ 关键词「{keyword}」搜索超时，跳过")
        return False
    except Exception as e:
        print(f"❌ 关键词「{keyword}」URL搜索失败：{str(e)[:30]}，跳过")
        return False


def clean_filename(filename):
    """清理图片文件名"""
    emoji_pattern = re.compile("["
                               u"\U0001F600-\U0001F64F" u"\U0001F300-\U0001F5FF" u"\U0001F680-\U0001F6FF" u"\U0001F1E0-\U0001F1FF"
                               "]+", flags=re.UNICODE)
    filename = emoji_pattern.sub(r'', filename)
    illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for c in illegal_chars:
        filename = filename.replace(c, '_')
    return filename


def download_image_no_cookie(img_url, save_path):
    """无Cookie下载图片"""
    try:
        if "video" in img_url or "gif" in img_url or "avatar" in img_url:
            return None
        if not img_url.startswith("http"):
            img_url = f"https:{img_url}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
            "Referer": "https://s.weibo.com/"
        }
        response = requests.get(img_url, headers=headers, timeout=15, stream=True)
        if response.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            return save_path
        else:
            return None
    except Exception as e:
        return None


def get_full_content(driver, card):
    """提取完整微博内容"""
    try:
        content_container = card.find_element(By.CSS_SELECTOR, "div[class*='content']")
        # 展开长文本
        try:
            expand_btn = card.find_element(By.XPATH,
                                           ".//a[contains(text(), '展开') or @action-type='feed_list_expand']")
            if expand_btn.is_displayed():
                driver.execute_script("arguments[0].click();", expand_btn)
                time.sleep(1)
        except:
            pass
        # 提取文本
        text_elems = content_container.find_elements(By.XPATH, ".//*[self::p or self::span or self::div]")
        content = " ".join([elem.text.strip() for elem in text_elems if elem.text.strip() and "展开" not in elem.text])
        return content
    except Exception as e:
        return ""


def extract_publisher_info(card):
    """提取发布者信息"""
    try:
        name_elem = card.find_element(By.CSS_SELECTOR, "a.name")
        user_name = name_elem.text.strip()
        href = name_elem.get_attribute("href")
        publisher_id = href.split("/")[-1].split("?")[0] if href else user_name
        return publisher_id, user_name
    except:
        return "未知", "未知"


def extract_publish_time(card):
    """提取发布时间"""
    try:
        time_elem = card.find_element(By.XPATH, ".//a[contains(@suda-data, 'wb_time')]")
        return time_elem.text.strip()
    except:
        return "未知"


def extract_like_count(card):
    """提取点赞数"""
    try:
        like_elem = card.find_element(By.CSS_SELECTOR, "span.woo-like-count")
        return int(like_elem.text.strip()) if like_elem.text.strip().isdigit() else 0
    except:
        return 0


def extract_forward_count(card):
    """提取转发数"""
    try:
        forward_elem = card.find_element(By.XPATH, ".//a[@action-type='feed_list_forward']")
        forward_num = re.findall(r'\d+', forward_elem.text.strip())
        return int(forward_num[-1]) if forward_num else 0
    except:
        return 0


def extract_comment_count(card):
    """提取评论数"""
    try:
        comment_elem = card.find_element(By.XPATH, ".//a[@action-type='feed_list_comment']")
        comment_num = re.findall(r'\d+', comment_elem.text.strip())
        return int(comment_num[-1]) if comment_num else 0
    except:
        return 0


def get_weibo_data(driver, keyword, page_num):
    """采集单页数据"""
    data = []
    time.sleep(WAIT_TIME)
    try:
        cards = driver.find_elements(By.CSS_SELECTOR,
                                     "div.card-wrap, div.feed-item, div[class*='wbpro-feed-item']"
                                     )
        if not cards:
            print(f"⚠️ 关键词「{keyword}」第{page_num}页无内容，停止采集")
            return []

        print(f"🔍 关键词「{keyword}」第{page_num}页找到 {len(cards)} 条微博")

        for card_idx, card in enumerate(cards, 1):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card)
                time.sleep(1)

                # 提取核心数据
                content = get_full_content(driver, card)
                publisher_id, user_name = extract_publisher_info(card)
                publish_time = extract_publish_time(card)
                like_count = extract_like_count(card)
                forward_count = extract_forward_count(card)
                comment_count = extract_comment_count(card)

                # 下载图片
                img_paths = []
                img_elems = card.find_elements(By.CSS_SELECTOR,
                                               "img[src*='.jpg'], img[src*='.png'], img[data-src*='.jpg']")
                for img_idx, img_elem in enumerate(img_elems):
                    img_url = img_elem.get_attribute("src") or img_elem.get_attribute("data-src")
                    if not img_url or "avatar" in img_url:
                        continue
                    base_name = f"{keyword}_page{page_num}_card{card_idx}_img{img_idx}"
                    img_save_path = os.path.join(IMG_SAVE_DIR, clean_filename(base_name) + ".jpg")
                    img_path = download_image_no_cookie(img_url, img_save_path)
                    if img_path:
                        img_paths.append(img_path)

                # 组装数据
                weibo_item = {
                    "检索关键词": keyword,
                    "页码": page_num,
                    "发布时间": publish_time,
                    "发布者ID": publisher_id,
                    "用户名称": user_name,
                    "点赞数": like_count,
                    "转发数": forward_count,
                    "评论数": comment_count,
                    "微博内容（完整）": content,
                    "图片路径": ";".join(img_paths) if img_paths else "无图片",
                    "采集时间": time.strftime("%Y-%m-%d %H:%M:%S")
                }
                data.append(weibo_item)
                print(f"  ✅ 第{card_idx}条：{user_name} | 点赞{like_count} | 转发{forward_count} | 评论{comment_count}")

            except StaleElementReferenceException:
                print(f"  ⚠️ 第{card_idx}条元素失效，跳过")
                continue
            except Exception as e:
                print(f"  ⚠️ 第{card_idx}条解析失败：{str(e)[:30]}")
                continue
    except Exception as e:
        print(f"❌ 第{page_num}页采集失败：{e}")
    return data


def turn_to_next_page(driver, keyword, current_page):
    """翻页（URL方式翻页，更稳定）"""
    try:
        # 方式1：URL翻页（优先）
        encoded_keyword = quote(keyword, encoding='utf-8')
        next_page_url = f"https://s.weibo.com/weibo?q={encoded_keyword}&page={current_page + 1}&Refer=STopic_weibo"
        driver.get(next_page_url)
        time.sleep(WAIT_TIME)

        # 检查是否有内容
        try:
            driver.find_element(By.CSS_SELECTOR, "div.card-wrap, div.feed-item")
            return True
        except:
            # 方式2：备用点击下一页按钮
            next_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '下一页') or @class='nextpage']"))
            )
            if "disabled" in next_btn.get_attribute("class"):
                return False
            driver.execute_script("arguments[0].click();", next_btn)
            time.sleep(WAIT_TIME)
            return True
    except:
        print("⚠️ 已到最后一页/无下一页，停止采集")
        return False


def save_data(all_data):
    """保存数据到Excel（解决超限 + 增强去重）"""
    if not all_data:
        print("💾 暂无数据可保存")
        return

    # 1. 转换为DataFrame并增强去重
    df = pd.DataFrame(all_data)
    # 增强去重：按【关键词+发布时间+用户名称+内容前50字】去重（避免长文本截断导致的重复）
    df["内容指纹"] = df["微博内容（完整）"].str.slice(0, 50).fillna("")  # 取前50字做指纹，处理空值
    df = df.drop_duplicates(
        subset=["检索关键词", "发布时间", "用户名称", "内容指纹"],
        keep="first"
    )
    df = df.drop(columns=["内容指纹"])  # 删除临时的内容指纹列

    # 2. 追加已有数据（如果文件存在）
    if os.path.exists(SAVE_EXCEL):
        try:
            existing_df = pd.read_excel(SAVE_EXCEL)
            # 已有数据也做一次去重，避免跨批次重复
            existing_df["内容指纹"] = existing_df["微博内容（完整）"].str.slice(0, 50).fillna("")
            existing_df = existing_df.drop_duplicates(
                subset=["检索关键词", "发布时间", "用户名称", "内容指纹"],
                keep="first"
            )
            existing_df = existing_df.drop(columns=["内容指纹"])

            # 合并新数据和已有数据
            df = pd.concat([existing_df, df], ignore_index=True)
            # 合并后再去重一次，确保绝对无重复
            df["内容指纹"] = df["微博内容（完整）"].str.slice(0, 50).fillna("")
            df = df.drop_duplicates(
                subset=["检索关键词", "发布时间", "用户名称", "内容指纹"],
                keep="first"
            )
            df = df.drop(columns=["内容指纹"])
        except Exception as e:
            print(f"⚠️ 读取已有Excel失败，将覆盖保存：{str(e)[:30]}")

    total_rows = len(df)
    print(f"📊 去重后总数据量：{total_rows} 行")

    # 3. 处理Excel行数超限（自动拆分）
    if total_rows <= MAX_EXCEL_ROWS:
        # 数据量未超限，直接保存
        df.to_excel(SAVE_EXCEL, index=False)
        print(f"💾 成功保存 {total_rows} 条数据到：{os.path.abspath(SAVE_EXCEL)}")
    else:
        # 数据量超限，自动拆分成多个文件
        base_name, ext = os.path.splitext(SAVE_EXCEL)
        file_index = 1
        for start in range(0, total_rows, MAX_EXCEL_ROWS):
            end = min(start + MAX_EXCEL_ROWS, total_rows)
            df_slice = df.iloc[start:end]
            split_file_name = f"{base_name}_part{file_index}{ext}"
            df_slice.to_excel(split_file_name, index=False)
            print(f"💾 拆分保存：{split_file_name}（{end - start} 行）")
            file_index += 1


def main():
    print("=" * 60)
    print("      微博全自动采集工具（URL搜索版）")
    print("=" * 60)

    # 1. 初始化浏览器
    driver = init_driver()
    # 先打开微博首页用于登录
    driver.get("https://weibo.com/")
    input("\n👉 请在浏览器中完成微博登录，登录后按回车继续...")

    # 2. 读取关键词
    keywords = read_keywords_from_excel(KEYWORD_EXCEL_PATH)

    # 3. 批量采集
    all_collect_data = []
    for idx, keyword in enumerate(keywords, 1):
        print(f"\n{'=' * 20} 开始采集第{idx}/{len(keywords)}个关键词：{keyword} {'=' * 20}")

        # URL方式直接搜索
        if not direct_search_by_url(driver, keyword):
            continue

        # 采集指定页数
        current_page = 1
        has_data = False
        while current_page <= COLLECT_PAGES_PER_KEYWORD:
            print(f"\n📖 采集「{keyword}」第{current_page}/{COLLECT_PAGES_PER_KEYWORD}页")
            page_data = get_weibo_data(driver, keyword, current_page)

            if page_data:
                all_collect_data.extend(page_data)
                save_data(all_collect_data)  # 每采集一页保存一次（保留原有逻辑）
                has_data = True
            else:
                break

            # 翻页（URL方式）
            if current_page < COLLECT_PAGES_PER_KEYWORD:
                if not turn_to_next_page(driver, keyword, current_page):
                    break
            current_page += 1

        if has_data:
            print(f"✅ 关键词「{keyword}」共采集 {current_page - 1} 页数据")
        else:
            print(f"⚠️ 关键词「{keyword}」未采集到有效数据")

    # 4. 收尾：最终保存一次（去重+拆分）
    save_data(all_collect_data)
    print("\n🎉 所有关键词采集完成！")
    print(f"🖼️ 图片保存路径：{os.path.abspath(IMG_SAVE_DIR)}")
    print(f"📝 最终结果文件：{os.path.abspath(SAVE_EXCEL)}（超限会自动拆分）")
    driver.quit()


if __name__ == "__main__":
    main()