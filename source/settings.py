import xml.dom.minidom
import codecs


def getsettings(tasksvariable, breaksvariable, workersminimum, breakslength, employees, version, excell_templates, excel_selected_variable):
    try:
        domtree = xml.dom.minidom.parse('settings.xml')
    except:
        domtree = xml.dom.minidom.Document()
        settings = domtree.createElement('settings')
        settings.setAttribute('version', version)

        new_task = domtree.createElement('task')
        new_task.setAttribute('id', '1')
        name = domtree.createElement('name')
        name.appendChild(domtree.createTextNode('Kassa'))
        new_task.appendChild(name)
        color = domtree.createElement('color')
        color.appendChild(domtree.createTextNode('e4b803'))
        new_task.appendChild(color)
        auto_generate = domtree.createElement('auto_generate')
        auto_generate.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(auto_generate)
        default_certified = domtree.createElement('default_certified')
        default_certified.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(default_certified)
        schedule_length = domtree.createElement('schedule_length')
        schedule_length.appendChild(domtree.createTextNode('780'))
        new_task.appendChild(schedule_length)
        schedule_max_times = domtree.createElement('schedule_max_times')
        schedule_max_times.appendChild(domtree.createTextNode('10'))
        new_task.appendChild(schedule_max_times)
        settings.appendChild(new_task)

        new_task = domtree.createElement('task')
        new_task.setAttribute('id', '2')
        name = domtree.createElement('name')
        name.appendChild(domtree.createTextNode('Rast'))
        new_task.appendChild(name)
        color = domtree.createElement('color')
        color.appendChild(domtree.createTextNode('000000'))
        new_task.appendChild(color)
        auto_generate = domtree.createElement('auto_generate')
        auto_generate.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(auto_generate)
        default_certified = domtree.createElement('default_certified')
        default_certified.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(default_certified)
        schedule_length = domtree.createElement('schedule_length')
        schedule_length.appendChild(domtree.createTextNode('780'))
        new_task.appendChild(schedule_length)
        schedule_max_times = domtree.createElement('schedule_max_times')
        schedule_max_times.appendChild(domtree.createTextNode('10'))
        new_task.appendChild(schedule_max_times)
        settings.appendChild(new_task)

        new_task = domtree.createElement('task')
        new_task.setAttribute('id', '3')
        name = domtree.createElement('name')
        name.appendChild(domtree.createTextNode('Arbetsledning'))
        new_task.appendChild(name)
        color = domtree.createElement('color')
        color.appendChild(domtree.createTextNode('1104b3'))
        new_task.appendChild(color)
        auto_generate = domtree.createElement('auto_generate')
        auto_generate.appendChild(domtree.createTextNode('false'))
        new_task.appendChild(auto_generate)
        default_certified = domtree.createElement('default_certified')
        default_certified.appendChild(domtree.createTextNode('false'))
        new_task.appendChild(default_certified)
        schedule_length = domtree.createElement('schedule_length')
        schedule_length.appendChild(domtree.createTextNode('780'))
        new_task.appendChild(schedule_length)
        schedule_max_times = domtree.createElement('schedule_max_times')
        schedule_max_times.appendChild(domtree.createTextNode('10'))
        new_task.appendChild(schedule_max_times)
        settings.appendChild(new_task)

        new_task = domtree.createElement('task')
        new_task.setAttribute('id', '4')
        name = domtree.createElement('name')
        name.appendChild(domtree.createTextNode('Förbutik'))
        new_task.appendChild(name)
        color = domtree.createElement('color')
        color.appendChild(domtree.createTextNode('780495'))
        new_task.appendChild(color)
        auto_generate = domtree.createElement('auto_generate')
        auto_generate.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(auto_generate)
        default_certified = domtree.createElement('default_certified')
        default_certified.appendChild(domtree.createTextNode('false'))
        new_task.appendChild(default_certified)
        schedule_length = domtree.createElement('schedule_length')
        schedule_length.appendChild(domtree.createTextNode('780'))
        new_task.appendChild(schedule_length)
        schedule_max_times = domtree.createElement('schedule_max_times')
        schedule_max_times.appendChild(domtree.createTextNode('10'))
        new_task.appendChild(schedule_max_times)
        settings.appendChild(new_task)

        new_task = domtree.createElement('task')
        new_task.setAttribute('id', '5')
        name = domtree.createElement('name')
        name.appendChild(domtree.createTextNode('SCO'))
        new_task.appendChild(name)
        color = domtree.createElement('color')
        color.appendChild(domtree.createTextNode('0c8d1c'))
        new_task.appendChild(color)
        auto_generate = domtree.createElement('auto_generate')
        auto_generate.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(auto_generate)
        default_certified = domtree.createElement('default_certified')
        default_certified.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(default_certified)
        schedule_length = domtree.createElement('schedule_length')
        schedule_length.appendChild(domtree.createTextNode('60'))
        new_task.appendChild(schedule_length)
        schedule_max_times = domtree.createElement('schedule_max_times')
        schedule_max_times.appendChild(domtree.createTextNode('3'))
        new_task.appendChild(schedule_max_times)
        settings.appendChild(new_task)

        new_task = domtree.createElement('task')
        new_task.setAttribute('id', '6')
        name = domtree.createElement('name')
        name.appendChild(domtree.createTextNode('PSS'))
        new_task.appendChild(name)
        color = domtree.createElement('color')
        color.appendChild(domtree.createTextNode('c86104'))
        new_task.appendChild(color)
        auto_generate = domtree.createElement('auto_generate')
        auto_generate.appendChild(domtree.createTextNode('false'))
        new_task.appendChild(auto_generate)
        default_certified = domtree.createElement('default_certified')
        default_certified.appendChild(domtree.createTextNode('true'))
        new_task.appendChild(default_certified)
        schedule_length = domtree.createElement('schedule_length')
        schedule_length.appendChild(domtree.createTextNode('60'))
        new_task.appendChild(schedule_length)
        schedule_max_times = domtree.createElement('schedule_max_times')
        schedule_max_times.appendChild(domtree.createTextNode('3'))
        new_task.appendChild(schedule_max_times)
        settings.appendChild(new_task)

        new = domtree.createElement('break')
        minimum = domtree.createElement('min')
        minimum.appendChild(domtree.createTextNode('60'))
        new.appendChild(minimum)
        maximum = domtree.createElement('max')
        maximum.appendChild(domtree.createTextNode('180'))
        new.appendChild(maximum)
        settings.appendChild(new)

        new = domtree.createElement('workers_minimum')
        data = domtree.createElement('h8')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h9')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h10')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h11')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h12')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h13')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h14')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h15')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h16')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h17')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h18')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h19')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        data = domtree.createElement('h20')
        data.appendChild(domtree.createTextNode('3'))
        new.appendChild(data)
        settings.appendChild(new)

        new = domtree.createElement('workingtime')
        data = domtree.createElement('h4')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h5')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h6')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h7')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h8')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h9')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h10')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h11')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('0'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h12')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h13')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h14')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        new.appendChild(data)
        data = domtree.createElement('h15')
        data2 = domtree.createElement('first_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('second_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        data2 = domtree.createElement('third_break')
        data2.appendChild(domtree.createTextNode('15'))
        data.appendChild(data2)
        data2 = domtree.createElement('forth_break')
        data2.appendChild(domtree.createTextNode('30'))
        data.appendChild(data2)
        new.appendChild(data)
        settings.appendChild(new)

        # excell footer template
        excell = domtree.createElement('excell')
        excell.setAttribute('id', '1')
        excell_title = domtree.createElement('title')
        excell_title.appendChild(domtree.createTextNode('Default'))
        excell.appendChild(excell_title)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'O:1')
        excell_cell.setAttribute('font', 'calibri')
        excell_cell.setAttribute('font_size', '14')
        excell_cell.setAttribute('font_style_bold', 'True')
        excell_cell.setAttribute('fg', '000000')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_left', 'thin')
        excell_cell.setAttribute('border_top', 'thin')
        excell_cell.appendChild(domtree.createTextNode('Tänk på att vi inväntar varandra ifrån rasterna'))
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'P:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Q:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'R:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'S:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'T:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'P:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'U:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'V:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'W:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'X:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Y:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Z:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AA:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AB:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AC:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AD:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AE:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AF:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AG:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AH:1')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_top', 'thin')
        excell_cell.setAttribute('border_right', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'O:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_left', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'P:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Q:2')
        excell_cell.setAttribute('bg', 'FF0000')
        excell_cell.setAttribute('font', 'calibri')
        excell_cell.setAttribute('font_size', '14')
        excell_cell.setAttribute('font_style_bold', 'True')
        excell_cell.setAttribute('fg', '000000')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.appendChild(domtree.createTextNode('Den som kommer tillbaka ifrån sin rast'))
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'R:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'S:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'T:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'U:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'V:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'W:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'X:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Y:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Z:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AA:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AB:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AC:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AD:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AE:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AF:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AG:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AH:2')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_right', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'O:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_left', 'thin')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'P:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Q:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'R:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'S:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell_cell.setAttribute('font', 'calibri')
        excell_cell.setAttribute('font_size', '14')
        excell_cell.setAttribute('font_style_bold', 'True')
        excell_cell.setAttribute('fg', '000000')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.appendChild(domtree.createTextNode('löser den som ska gå på rast.'))
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'T:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'U:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'V:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'W:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'X:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Y:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'Z:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AA:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AB:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AC:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AD:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AE:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AF:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AG:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell.appendChild(excell_cell)
        excell_cell = domtree.createElement('cell')
        excell_cell.setAttribute('id', 'AH:3')
        excell_cell.setAttribute('bg', 'FFFF0000')
        excell_cell.setAttribute('border_bottom', 'thin')
        excell_cell.setAttribute('border_right', 'thin')
        excell.appendChild(excell_cell)
        settings.appendChild(excell)

        excel_selected = domtree.createElement('excel_selected')
        excel_selected.setAttribute('id', '1')
        settings.appendChild(excel_selected)

        domtree.appendChild(settings)

        domtree.writexml(codecs.open('settings.xml', "w", "utf-8"), encoding="utf-8")
    settings = domtree.documentElement

    v = settings.getAttribute('version')

    # upgrade to version 0.1.1
    if not v == '0.1.0':
        settings.setAttribute('version', version)

        domtree.writexml(codecs.open('settings.xml', "w", "utf-8"), encoding="utf-8")

    tasks = settings.getElementsByTagName('task')
    for task in tasks:
        temp = []
        temp.append(task.getAttribute('id'))
        temp.append(task.getElementsByTagName('name')[0].childNodes[0].nodeValue)
        temp.append(task.getElementsByTagName('color')[0].childNodes[0].nodeValue)
        temp.append(
            bool(eval(task.getElementsByTagName('auto_generate')[0].childNodes[0].nodeValue.lower().capitalize())))
        temp.append(
            bool(eval(task.getElementsByTagName('default_certified')[0].childNodes[0].nodeValue.lower().capitalize())))
        temp.append(task.getElementsByTagName('schedule_length')[0].childNodes[0].nodeValue)
        temp.append(task.getElementsByTagName('schedule_max_times')[0].childNodes[0].nodeValue)
        tasksvariable.append(temp)

    breaks = settings.getElementsByTagName('break')
    for b in breaks:
        temp = []
        temp.append(b.getElementsByTagName('min')[0].childNodes[0].nodeValue)
        temp.append(b.getElementsByTagName('max')[0].childNodes[0].nodeValue)
        breaksvariable.append(temp)

    w = settings.getElementsByTagName('workers_minimum')
    for i in range(13):
        workersminimum.append(w[0].getElementsByTagName('h' + str(i + 8))[0].childNodes[0].nodeValue)

    workingtime = settings.getElementsByTagName('workingtime')
    for i in range(12):
        c = []
        b = workingtime[0].getElementsByTagName('h' + str(i + 4))
        c.append(b[0].getElementsByTagName('first_break')[0].childNodes[0].nodeValue)
        c.append(b[0].getElementsByTagName('second_break')[0].childNodes[0].nodeValue)
        c.append(b[0].getElementsByTagName('third_break')[0].childNodes[0].nodeValue)
        c.append(b[0].getElementsByTagName('forth_break')[0].childNodes[0].nodeValue)
        breakslength.append(c)

    employee = settings.getElementsByTagName('employee')
    for e in employee:
        name = e.getElementsByTagName('name')[0].childNodes[0].nodeValue
        default_task = e.getElementsByTagName('default_task')[0].childNodes[0].nodeValue
        temp = []
        tasks = e.getElementsByTagName('task_settings')
        for task in tasks:
            temp.append([task.getAttribute('id'), task.getElementsByTagName('certified')[0].childNodes[0].nodeValue])
        employees.append([name, temp, default_task])

    # load excell template
    excell_templates['0'] = ['Empty', []]
    excells = settings.getElementsByTagName('excell')
    for excell in excells:
        data = []
        cells = excell.getElementsByTagName('cell')
        for cell in cells:
            temp = {}
            temp['id'] = cell.getAttribute('id')
            temp['font'] = cell.getAttribute('font')
            temp['font_size'] = cell.getAttribute('font_size')
            temp['font_style_bold'] = cell.getAttribute('font_style_bold')
            temp['font_style_italic'] = cell.getAttribute('font_style_italic')
            temp['font_style_underline'] = cell.getAttribute('font_style_underline')
            temp['fg'] = cell.getAttribute('fg')
            temp['bg'] = cell.getAttribute('bg')
            temp['border_left'] = cell.getAttribute('border_left')
            temp['border_right'] = cell.getAttribute('border_right')
            temp['border_top'] = cell.getAttribute('border_top')
            temp['border_bottom'] = cell.getAttribute('border_bottom')
            temp['border_left_color'] = cell.getAttribute('border_left_color')
            temp['border_right_color'] = cell.getAttribute('border_right_color')
            temp['border_top_color'] = cell.getAttribute('border_top_color')
            temp['border_bottom_color'] = cell.getAttribute('border_bottom_color')
            if cell.childNodes:
                temp['text'] = cell.childNodes[0].nodeValue
            else:
                temp['text'] = ''
            data.append(temp)
        title = excell.getElementsByTagName('title')
        excell_templates[excell.getAttribute('id')] = [title[0].childNodes[0].nodeValue, data]
    excel_selected = settings.getElementsByTagName('excel_selected')[0]
    excel_selected_variable[0] = excel_selected.getAttribute('id')
