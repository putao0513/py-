from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# === 工具函数：记录耗时 ===
def log_time(step_name, start_time):
    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    print(f"⏱️ {step_name} 完成，耗时 {minutes}分{seconds}秒")

script_start = time.time()

# === 1️⃣ 启动浏览器 ===
step_start = time.time()
options = webdriver.EdgeOptions()
options.add_argument(r"user-data-dir=D:\pyproject\auto\edge_user_data")  # 使用固定用户数据目录
options.add_argument("--start-maximized")

service = Service(r"D:\pyproject\auto\edgedriver_win64\msedgedriver.exe")
driver = webdriver.Edge(service=service, options=options)
log_time("启动浏览器", step_start)

# === 2️⃣ 打开页面 ===
step_start = time.time()
url = "https://fssc.ysservice.com.cn/boe/GL_BOE?boeTypeCode=AMORTIZATION_STATEMENT"
driver.get(url)
time.sleep(3)
log_time("打开页面", step_start)

# === 3️⃣ 判断是否需要登录 ===
step_start = time.time()
try:
    username_input = driver.find_element(By.ID, "username")
    password_input = driver.find_element(By.ID, "selfpwd")
    print("🔐 检测到登录输入框，开始登录...")
    username_input.send_keys("zhangqiang28")
    password_input.send_keys("@cifi.com.cn")
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "button.layui-btn").click()
    time.sleep(5)
except:
    print("✅ 已检测到已登录状态，无需输入账号密码")
log_time("登录检测/输入", step_start)

# === 4️⃣ 填写业务类型 ===
step_start = time.time()
biz_type_input = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.el-input__inner[placeholder="请选择，可输入关键字查询"]'))
)
biz_type_input.clear()
biz_type_input.send_keys("退款及其他类代收代付")
time.sleep(1)
biz_type_input.send_keys(Keys.ENTER)
time.sleep(2)
log_time("填写业务类型", step_start)

# === 5️⃣ 填写费用支付公司（唯一编号 + 点击候选行） ===
step_start = time.time()
le_input = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '#boeHeader\\.0\\.leId input.el-input__inner'))
)
le_input.click()
time.sleep(0.5)

unique_code = "A-756B-001"
le_input.send_keys(unique_code)
time.sleep(1)  # 等待下拉渲染

# 找到下拉列表中唯一匹配的行并点击
rows = driver.find_elements(By.CSS_SELECTOR, '.el-table__body-wrapper .el-table__row')
target_row = None
for row in rows:
    if unique_code in row.text:
        target_row = row
        break

if target_row:
    ActionChains(driver).move_to_element(target_row).click().perform()
    print("✅ 费用支付公司已选中:", unique_code)
else:
    print("⚠️ 未找到匹配行，尝试回车确认")
    le_input.send_keys(Keys.ENTER)
time.sleep(1)
log_time("填写费用支付公司", step_start)

# === 6️⃣ 上传表格 ===
step_start = time.time()
upload_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.upfile.inport-temp input[type="file"]'))
)
file_path = r"D:\费用分摊流程导入表格.xlsx"
upload_input.send_keys(file_path)
time.sleep(2)
print("✅ 已完成登录/表单填写/上传表格，浏览器保持开启")
log_time("上传表格", step_start)

# === 7️⃣ 点击明细第一行的删除按钮 ===
step_start = time.time()
try:
    time.sleep(1)  # 等待明细表格渲染
    delete_buttons = driver.find_elements(By.XPATH, '//button[contains(@class, "el-button--text") and contains(@class, "el-button--small")]/span[text()="删除"]/..')
    if delete_buttons:
        delete_buttons[0].click()
        print("✅ 已点击明细第一行删除按钮")
    else:
        print("⚠️ 没有找到删除按钮")
except Exception as e:
    print("❌ 删除操作失败:", e)
log_time("点击明细删除按钮", step_start)

# === 8️⃣ 保持浏览器开启等待手动检查 ===
total_elapsed = time.time() - script_start
m, s = divmod(int(total_elapsed), 60)
print(f"🏁 脚本全部完成，总耗时 {m}分{s}秒")

input("操作完成，浏览器保持开启，按 Enter 键退出脚本...")
