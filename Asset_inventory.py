import time

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyodbc
import datetime
import subprocess
from datetime import datetime
import os

"""
Имеется сайт, на котором нужно заполнять форму, хранящую данные о технике(конкретно ноутбук), короче говоря - инвентаризация.
Пункты для заполнения:
Сетевое имя ноутбука, ФИО того, у кого на активе, тип оборудования(ноутбук) Серийный номер, производитель, модель
полное описание ноутбука, дата заполнения, инвентарный номер, организация которой выдали ноутбук, месторасположение, 
комментарий к месторасположению, статус оборудования( на складе или выдан)
"""

"""
дабы облегчить это заполнение, написал код, который:
1) проходит по ссылке на сайт для заполнения, нужно будет ввести логин-пароль от него
2) запросит некоторую информацию, которую не автоматизировать: данные заполняющего( который и будет ответственным за актив, если
на склад)
то есть, если техника на склад уходит, ответственный тот, кто заполняет, если же техника выдается, то в процессе запрашивается 
ФИО того, кому передается ноутбук
3) спросит куда уходит техника - склад или на выдачу
4) инвентарный номер
5) Производитель и модель ноутбука( поскольку информация берется из таблицы ms access, немного проблематично извлечь эти два 
параметра из базы, так как база заполняется другими людьми в хаотичной форме
"""

class NotebookStock:

    #с вызовом конструктора добавим аргументы,чтобы браузер запускался корректно
    def __init__(self, user_p, status_p):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        # options.add_argument("--window-size=1920,1080")
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument("--disable-extensions")
        options.add_argument("--proxy-server='direct://'")
        options.add_argument("--proxy-bypass-list=*")
        # options.add_argument("--headless=new")
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--blink-settings=imagesEnabled=false')

        self.driver = webdriver.Chrome(options=options)
        self.user = user_p
        self.status = status_p
        self.other_user = ""

    def filling_form(self):   #заполняем форму, целая функция на все телодвижения, внутри разные if конструкции
        self.driver.get("https://its.terralink-global.com/AssetManagement/HardwareAsset/New")
        self.driver.implicitly_wait(30)


        self.driver.find_element(By.NAME, "Name").send_keys(self.get_sn_from_cmd("hostname"), Keys.ENTER)

        if self.status == "На складе":
            self.driver.find_element(By.NAME, "Target_HardwareAssetHasPrimaryUser").send_keys(self.user, Keys.ENTER)
        elif self.status == "Выдан":
            self.other_user = str(input("""Введите фамилию и имя того, на кого записываем оборудование. Формат - 
            Eleusizov, Muhammed\n обязательно как в базе Терралинка"""))
            self.driver.find_element(By.NAME, "Target_HardwareAssetHasPrimaryUser").send_keys(self.other_user, Keys.ENTER)


        self.driver.find_element(By.NAME, "SerialNumber").send_keys(self.get_sn_from_cmd("wmic bios get serialnumber"),
                                                                    Keys.ENTER)
        asset = self.get_item_from_accdb()

        self.driver.find_element(By.NAME, "Manufacturer").send_keys(str(input("Введите производителя\n")), Keys.ENTER)
        self.driver.find_element(By.NAME, "Model").send_keys(str(input("Введите модель\n")), Keys.ENTER)

        self.driver.find_element(By.XPATH, '/html/body/div[7]/div[3]/div/div[1]/div[2]/div/div/div\
        /div[1]/div[1]/div/div[2]/div[6]/div[2]/div/div/span/span/input').send_keys("Notebook", Keys.ENTER)

        self.driver.find_element(By.XPATH, '/html/body/div[7]/div[3]/div/div[1]/div[2]/div/div/div\
        /div[1]/div[1]/div/div[2]/div[6]/div[1]/div/div/span/span/input').send_keys(self.status, Keys.ENTER)

        self.driver.find_element(By.NAME, "Description").send_keys(asset[2] + "   " + str(asset[3]) + " тг", Keys.ENTER)

        if self.user == "Eleusizov, Muhammed":
            self.driver.find_element(By.NAME, "Target_HardwareAssetHasLocation").send_keys("KZ-NST-Office", Keys.ENTER)
            self.driver.implicitly_wait(10)
            self.driver.find_element(By.NAME, "LocationDetails").send_keys("офис г. Астана", Keys.ENTER)
        else:
            self.driver.find_element(By.NAME, "Target_HardwareAssetHasLocation").send_keys("KZ-ALA-Office", Keys.ENTER)
            self.driver.implicitly_wait(10)
            self.driver.find_element(By.NAME, "LocationDetails").send_keys("офис г. Алматы", Keys.ENTER)

        self.driver.find_element(By.NAME, "Target_HardwareAssetHasOrganization").send_keys("Terralnk Kazakhstan", Keys.ENTER)

        dt_string = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.driver.find_element(By.NAME, 'ManualInventoryDate').send_keys(dt_string, Keys.ENTER)

        finance_tab = self.driver.find_element(By.XPATH,
                                               '/html/body/div[7]/div[3]/div/div[1]/div[2]/div/div/ul/li[5]/a')
        ActionChains(self.driver).move_to_element(finance_tab).click().perform()

        if self.status == "На складе":
            self.driver.find_element(By.NAME, "OwnedBy").send_keys(self.user, Keys.ENTER)
        elif self.status == "Выдан":
            self.driver.find_element(By.NAME, "OwnedBy").send_keys(self.other_user, Keys.ENTER)

        self.driver.find_element(By.NAME, "Kaz_Inventory_Number").send_keys(asset[1], Keys.ENTER)
        time.sleep(1000)

        '''
        print("""Введите дату получения из вкладки "Финансы" в формате - 18/07/2023 00:00 """)
        rec_date_inp = input()
        self.driver.find_element(By.NAME, "ReceivedDate").send_keys(rec_date_inp, Keys.ENTER)'''

    def get_sn_from_cmd(self, command): #функция для получения некоторых данных с ноутбука для заполнения формы
        if command == "hostname":
            return subprocess.check_output(command, shell=True, text=True)
        else:
            try:
                # Запускаем команду wmic и передаем ее как строку для исполнения
                res = []
                result = subprocess.check_output(command, shell=True, text=True)
                lines = result.splitlines()

                # Выводим каждую строку отдельно
                for line in lines:
                    res.append(line)

                return res[2]       #возвращаем нужную строку

            except subprocess.CalledProcessError as e:
                print("Ошибка выполнения команды:", e)

    def get_item_from_accdb(self):  #связь с ms access по сетевому пути
        connection_string = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=\\s-fs03-ala\akt\Акт приема-перадачи.accdb;'       #сетевой путь к БД
        )

        try:
            # Установка соединения
            conn = pyodbc.connect(connection_string)

            # Создание курсора
            cursor = conn.cursor()

            inventory_value = str(input("Введите инвентарный номер\n"))
            column_name = "[Инвентарный номер]"
            # Параметризованный SQL-запрос с использованием плейсхолдера '?'
            query = f"SELECT * FROM ОС WHERE {column_name} = ?"   #поиск по столбцу "Инвентарный номер" на совпадение

            # Выполнение запроса с передачей значения параметра в виде кортежа
            cursor.execute(query, (inventory_value,))

            item_info = cursor.fetchone()

            # Проверка, найдена ли строка
            if item_info:
                print("Найденная строка:")
                print(item_info)
                return item_info
            else:
                print("Строка с таким инвентарным номером не найдена.")

            # Закрытие курсора и соединения
            cursor.close()
            conn.close()

        except pyodbc.Error as e:
            print('Ошибка:', e)


    # для корректной связи с ms access, требуется установка программы ниже, установка пройдет в тихом режиме
    def install_accdb_engine(self):
        program_name = "Microsoft Access database engine 2010 (English)"
        try:
            # Выполнение команды wmic для получения списка установленных программ
            result = subprocess.run(["wmic", "product", "get", "Name","/format:list"], capture_output=True, text=True)

            # Проверяем установлена ли программа
            if program_name not in result.stdout:
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

                current_directory = os.getcwd()  #получаем директорию и меняем на Desktop, следовательно, установщик нужно закинуть
                #в Desktop

                os.chdir(desktop_path)

                updated_directory = os.getcwd()

                try:
                    result = subprocess.run("AccessDatabaseEngine_X64.exe /quiet /norestart", shell=True, text=True)
                    if result:
                        print("AccessDatabaseEngine установлен")
                except subprocess.CalledProcessError as e:
                    print("Ошибка выполнения команды:", e)
            else:
                return False
        except Exception as e:
            print("Ошибка:", e)
            return False

def main():
    user_that_fills = 0
    users_dict = {1: "Eleusizov, Muhammed", 2: "Dzhangashkarov, Erlan", 3: "Rakishev, Yerasyl"}
    asset_status = {1: "На складе", 2: "Выдан"}
    status = 0

    while user_that_fills not in users_dict:
        try:
            print("Кто будет заполнять, 1 - Мухаммед, 2 - Ерлан, 3 - Ерасыл? Ответ числом")
            user_that_fills = int(input())
        except ValueError:
            print("Что-то пошло не так. Давайте заново")

    while status not in asset_status:
        try:
            print("Что вы хотите сделать с оборудованием: 1 - на складе, 2 - выдать. Ответ числом")
            status = int(input())
        except ValueError:
            print("Что-то пошло не так. Давайте заново")

    ns = NotebookStock(users_dict[user_that_fills], asset_status[status])
    ns.install_accdb_engine()
    ns.filling_form()


if __name__ == '__main__':
    main()

