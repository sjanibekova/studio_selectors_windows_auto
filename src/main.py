import json
import time
from io import BytesIO
import win32api
import pywinauto
from ctypes.wintypes import tagPOINT
from pprint import pprint as pp

import win32gui
from win32api import GetSystemMetrics
from UIDesktop import UIO_Highlight, UIOEI_Convert_UIOInfo, UIOSelector_Get_UIOList, UIOSelector_Get_UIO
import base64
import mouse
import keyboard
from comtypes import COMError


mDefaultPywinautoBackend = "uia"


def get_image_of_element(wrapper):
    """
        Encode image to base64

    :param wrapper:
    :return: image encoded to base64
    """
    image = wrapper.capture_as_image()
    img = image
    im_file = BytesIO()
    img.save(im_file, format="JPEG")
    # img.show()
    im_bytes = im_file.getvalue()  # im_bytes: image in binary format.
    im_b64 = base64.b64encode(im_bytes)
    return im_b64


# Get list of top level
def backend_str_get_top_level_list_uio_info(in_backend=mDefaultPywinautoBackend):
    """ Gets all current active windows
    :param inBackend: uia or win32
    :return: lResultList2 - List with information about window (title, classname, control_type)
    """
    # Получить список объектов
    l_result_list = pywinauto.findwindows.find_elements(top_level_only=True, backend=in_backend)
    l_result_list2 = []
    for lI in l_result_list:
        lTempObjectInfo = lI
        l_result_list2.append(UIOEI_Convert_UIOInfo(lI, backend=in_backend))
    return l_result_list2


# print(BackendStr_GetTopLevelList_UIOInfo())


# получить активный елемент
def get_current_element_by_coordinates():
    """ Gets current mouse position and get control information
            about window
        :return: element - UIAElementInfo pywinauto - information about window
                 wrapper - UIAWRapper  pywindauto - give access to additional manipilation of window
        """
    # получить текущие координаты
    # com_type errors
    try:
        (x, y) = win32api.GetCursorPos()
        # управление элементов через координаты
        elem = pywinauto.uia_defines.IUIA().iuia.ElementFromPoint(tagPOINT(x, y))
        # чтобы получить спец  - нужно обернуть в ElementInfo
        element = pywinauto.uia_element_info.UIAElementInfo(elem)
        wrapper = pywinauto.controls.uiawrapper.UIAWrapper(element)
        return element, wrapper
    except BaseException as e:
        return [None, None]



def get_full_specification_of_element(element):
    """ Add full information into dictionary format
    :arg: element - pywinauto Element
    :returns: full_element_info in dict format
    """
    full_element_info = {"control_id": element.control_id, "enabled": element.enabled, "handle": element.handle,
                         "title": element.name, "process_id": element.process_id, "rectangle": str(element.rectangle),
                         "rich_text": element.rich_text, "control_type": element.control_type,
                         # Не уверена что правильно
                         "runtime_id": element.runtime_id, "class_name": element.class_name, "children": [{"ctrl_index":(index, element.name)} for index, element in enumerate(element.children())]}
    return full_element_info


def reindex_elements_in_tree(result):
    """
    Change the level numbers from parent to chil
    Before it was from children to parent
    :param result: dict format of element specification
    :return: result_list reindexed
    """
    result_list = [values for index, values  in result.items()]
    # pp(result_list)
    # Убираем родителя для всех окон(рабочий стол) и оставляем спец окна
    reversed_result = list(reversed(result_list))[1:]
    result_list = [{"level " + str(index): item}
                   for index, item in enumerate(reversed_result)]
    return result_list


def get_element_tree_structure(element, wrapper):
    """ Get get element and return all its parents and convert info
        about element according to specification
    :param element:
    :return: dict_of_tree contains info from child to parent
    """
    dict_of_tree = {}
    parent_level_index = 0
    element_parent = element
    # print(wrapper.top_level_parent(), element_parent)
    while element_parent:
        # уровень от 0 - (уровень елемента) и -1 для родителя child_level: 0 parent level: -1 ... parent of parent: -2
        level = parent_level_index - 1
        # Добавляем информацию о родителях по одному

        dict_of_tree[level] = get_full_specification_of_element(element_parent)
            # обратись  к родителю
        element_parent = element_parent.parent
        parent_level_index = parent_level_index - 1

        # getFullInfoAboutElement(element)'
    return dict_of_tree


#has_depth(root, depth)
# Return True if element has particular depth level relative to the root
#
# def get_all_windows()

def search_selector_run(keymap=None):
    """
        Program search elements on window until ctrl is not pushed
        Then return full selector path of info
    """
    wrapper_manager = []
    if keymap:
        button = keymap["mouse_button"]
        hothey = keymap["hotkey"]
        # Todo если мышка то левая или правая кнопка
        # Todo hot key - передать код клавиатуры
    else:
        # Default hot key
        button = "left"
        hothey = "ctrl"

    lFlagLoop = True
    # Пока зажат ctrl продолжай цикл
    while lFlagLoop:
        # Получай элемент по координанам
        try:
            element, wrapper = get_current_element_by_coordinates()

            # проверка зажата ли правая кнопка мыши или ctrl
            if mouse.is_pressed(button=button) and keyboard.is_pressed(hothey):
                result = get_element_tree_structure(element, wrapper)

                final_result = reindex_elements_in_tree(result)
                element_screenshot = get_image_of_element(wrapper)
                final_result.append({"backend": "uia"})
                final_result.append({"image_base_64_code": element_screenshot.decode('utf-8')})
                lFlagLoop = False
            else:
                # очистить холст
                '''
                https://stackoverflow.com/questions/7207309/how-to-run-functions-in-parallel
                https://stackoverflow.com/questions/18864859/python-executing-multiple-functions-simultaneously
                
                https://stackoverflow.com/questions/37305014/python-running-n-number-of-functions-in-parallel
                # Into class 
                '''

                time.sleep(0.3)
                element1, wrapper1 = get_current_element_by_coordinates()
                if wrapper1 == wrapper:
                    UIO_Highlight(wrapper)
                # очистить холст
                else:
                    clear_window_after_drawing()
        except COMError as e:
            print(e)
            continue
    pp(final_result)
    with open('selector_data.json', 'w') as f:
        json.dump(final_result, f)
    return final_result


def clear_window_after_drawing():
    hwnd = win32gui.WindowFromPoint((0, 0))
    monitor = (0, 0, GetSystemMetrics(0), GetSystemMetrics(1))
    win32gui.InvalidateRect(hwnd, monitor, True)




def main():
    level_tree = search_selector_run()
    # pp(level_tree)
    selector = []
    iteration = 0
    last_element_rectangle = ""
    max_level = 0

    for i in range(len(level_tree)):
        if "backend" == list(level_tree[i].keys())[0]:
            backend = level_tree[i]["backend"]
        if "level" in list(level_tree[i].keys())[0]:
            max_level = max(max_level, int(list(level_tree[i].keys())[0].replace("level ", "")))

    for i in range(len(level_tree)):
        if "level" in list(level_tree[i].keys())[0]:
            values_view = level_tree[i].values()
            value_iterator = iter(values_view)
            level_value = next(value_iterator)
            level_dict = {"class_name": level_value["class_name"], "control_type": level_value["control_type"],
                          "title": level_value["title"]}
            if list(level_tree[i].keys())[0] == "level 0":
                level_dict["backend"] = backend
                level_dict.pop("control_type")
            if i == max_level:
                last_element_rectangle = level_value["rectangle"]
            selector.append(level_dict)
    # Не забудь передать backend
    pp(selector)
    same_list = UIOSelector_Get_UIOList(selector)
    pp(same_list)

    if len(same_list) > 1:
        for i in range(len(same_list)):
                if str(same_list[i].rectangle()) == last_element_rectangle:
                    iteration = i
                    return same_list[i]
    return same_list[0]
'''
draw_outline(colour='green', thickness=2, fill=<MagicMock name='mock.win32defines.BS_NULL' id='140124673757368'>, rect=None)
'''
if __name__ == '__main__':
    main()
    # time.sleep(20)