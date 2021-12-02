import eel
from get_pic import *


@eel.expose
def do(filepath1):
    print(f'Выбран файлы: \n{filepath1 }')
    if main(filepath1):
        pass


def start():
    eel.init('web')  # Подключаем папку с интерфейсом
    eel.start('main.html', size=(700, 800))


if __name__ == '__main__':
    start()

