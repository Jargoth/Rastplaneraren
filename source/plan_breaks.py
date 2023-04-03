# This file contains everyting needed to plan breaks

import random
from tkinter import messagebox

def plan_breaks(generateoptions, person, breakslength, breaksvariable, workersminimum, tasksvariable, employees):
    # autogenerate breaks, tasks and priority

    # generate the breaks
    if generateoptions[0].get():
        breaks(person, breakslength, breaksvariable, workersminimum, tasksvariable)

    # generate the tasks
    if generateoptions[1].get():
        tasks(tasksvariable, person, employees)

    # generate the priorities
    if generateoptions[2].get():
        priorities(person)

    # if none is set
    if not generateoptions[0].get() and not generateoptions[1].get() and not generateoptions[2].get():
        messagebox.showwarning('Varning', 'Inget alternativ är valt.')


def breaks(person, breakslength, breaksvariable, workersminimum, tasksvariable):
    workersminimum_override = 0
    availableworkers = []
    for i in range(52):
        availableworkers.append([0, 0])
    for i in range(len(person)):
        for j in range(52):
            if person[i][4][j][1] != -1 and person[i][4][j][1] != 1:
                availableworkers[j][0] = availableworkers[j][0] + 1
            if person[i][4][j][1] == 1:
                availableworkers[j][1] = availableworkers[j][1] + 1
    for i in range(len(person)):
        if person[i][2].get():
            workingtime = person[i][2].get().split('-')
            starttime = workingtime[0].split(':')
            starttime = (int(starttime[0]) * 60) + int(starttime[1])
            endtime = workingtime[1].split(':')
            endtime = (int(endtime[0]) * 60) + int(endtime[1])
            totaltime = int((endtime - starttime) / 60)
            if (totaltime - 4) >= 0:  # if working long enough to get a break
                for breaknumber, bs in enumerate(breakslength[totaltime - 4]):
                    if int(bs):
                        for j in range(len(person[i][4])):
                            if person[i][4][j][1] == 1:
                                starttime = 8 * 60 + j * 15
                        bs = int(int(bs) / 15)
                        temp = (int(starttime + int(breaksvariable[0][1]) - int(breaksvariable[0][0])) / 15) - 32
                        notdone = 1
                        forward = False
                        numtries = 0
                        maxminbreak = 0
                        simultaneusbreaks = 1
                        while notdone:
                            ok = 0
                            for b in range(bs):
                                if availableworkers[int(temp) + b][0] > (
                                        int(workersminimum[int((temp + b) / 4)]) + workersminimum_override) and \
                                        availableworkers[int(temp)][1] == simultaneusbreaks - 1:
                                    ok = ok + 1
                            if ok == bs:

                                # calculate remaining working time including this break
                                remainingtime = 0
                                for remaining in range(52 - (int(temp) + 1)):
                                    if person[i][4][int(temp) + remaining][1] > -1:
                                        remainingtime = remainingtime + 1

                                # calculate min/max remaining time reqired
                                breaksremaining = 0
                                iteration = breaknumber
                                while iteration < 4:
                                    if int(breakslength[totaltime - 4][iteration]):
                                        breaksremaining = breaksremaining + 1
                                    iteration = iteration + 1
                                remainingtimemin = int(breaksremaining * int(breaksvariable[0][0]) / 15)
                                remainingtimemax = int(breaksremaining * int(breaksvariable[0][1]) / 15)
                                iteration = breaknumber
                                while iteration < 4:
                                    remainingtimemin = \
                                        remainingtimemin + int(int(breakslength[totaltime - 4][iteration]) / 15)
                                    remainingtimemax = \
                                        remainingtimemax + int(int(breakslength[totaltime - 4][iteration]) / 15)
                                    iteration = iteration + 1

                                # set the time offset to correct the break
                                if remainingtimemin <= remainingtime and remainingtimemax >= remainingtime:
                                    offsettime = 0
                                elif remainingtime < remainingtimemin:
                                    offsettime = remainingtime - remainingtimemin
                                elif remainingtime > remainingtimemax:
                                    offsettime = remainingtime - remainingtimemax
                                else:
                                    offsettime = 0

                                for b in range(bs):
                                    availableworkers[int(temp) + b + offsettime][0] = \
                                        availableworkers[int(temp) + offsettime][0] - 1
                                    availableworkers[int(temp) + b + offsettime][1] = \
                                        availableworkers[int(temp) + offsettime][1] + 1

                                    person[i][4][int(temp) + b + offsettime][0]['bg'] = f'#{tasksvariable[1][2]}'
                                    person[i][4][int(temp) + b + offsettime][1] = 1
                                notdone = 0

                            elif forward:
                                temp = (int(starttime + int(breaksvariable[0][1]) - int(
                                    breaksvariable[0][0])) / 15) - 32 + 1 + numtries
                                forward = False
                                numtries = numtries + 1
                                if temp - (int(starttime / 15) - 32) > int(int(breaksvariable[0][1]) / 15):
                                    maxminbreak = maxminbreak + 1
                            else:
                                temp = (int(starttime + int(breaksvariable[0][1]) - int(
                                    breaksvariable[0][0])) / 15) - 32 - 1 - numtries
                                forward = True
                                if temp - (int(starttime / 15) - 32) - (b / 15) - 1 < int(int(breaksvariable[0][0]) / 15):
                                    maxminbreak = maxminbreak + 1
                            if maxminbreak == 2:
                                simultaneusbreaks = simultaneusbreaks + 1
                                if simultaneusbreaks == 5:
                                    simultaneusbreaks = 1
                                    workersminimum_override = workersminimum_override - 1
                                numtries = 0
                                maxminbreak = 0
                                temp = \
                                    (int(starttime + int(breaksvariable[0][1]) - int(breaksvariable[0][0])) / 15) - 32


def tasks(tasksvariable, person, employees):
    for i in range(2, len(tasksvariable)):
        if tasksvariable[i][3]:
            nextemployee = 0
            workers = []
            for j, p in enumerate(person):
                if p[0].get():
                    for employee in employees:
                        if p[0].get().lower().capitalize() == employee[0]:
                            for emptask in employee[1]:
                                if emptask[0] == tasksvariable[i][0] and eval(emptask[1]):
                                    workers.append([j, 0])
            random.shuffle(workers)
            for quarter in range(52):
                done = False
                for p in person:
                    if p[4][quarter][1] == i:
                        done = True
                        break
                if not done:
                    for j in range(len(workers)):

                        # see if there is 4 free quarters next, including this quarter
                        num_free_quarter = 0
                        for free_quarter in range(4):
                            if (quarter + free_quarter) >= 52:
                                num_free_quarter = num_free_quarter + 1
                            else:
                                if person[workers[nextemployee][0]][4][quarter + free_quarter][1] == 0:
                                    num_free_quarter = num_free_quarter + 1
                                # see how much time to plan
                                required_quarter = 0
                                for p in person:
                                    if p[4][quarter + free_quarter][1] != i:
                                        required_quarter = required_quarter + 1
                                    else:
                                        break

                        timesplanned = 0
                        print(f'{i} {quarter} {employees[nextemployee][0]} {num_free_quarter} {required_quarter}')
                        if person[workers[nextemployee][0]][4][quarter + timesplanned][1] == 0 and \
                                workers[nextemployee][1] < int(tasksvariable[i][6]) and \
                                (num_free_quarter >= 4 or num_free_quarter >= required_quarter):
                            planned = False
                            while timesplanned < int(tasksvariable[i][5]) / 15 and \
                                    person[workers[nextemployee][0]][4][quarter + timesplanned][1] == 0:
                                if quarter + timesplanned < 52:
                                    if person[workers[nextemployee][0]][4][quarter + timesplanned][1] == 0:
                                        done = False
                                        for p in person:
                                            if p[4][quarter + timesplanned][1] == i:
                                                done = True
                                                break
                                        if not done:
                                            person[workers[nextemployee][0]][4][quarter + timesplanned][1] = i
                                            person[workers[nextemployee][0]][4][quarter + timesplanned][0][
                                                'bg'] = f'#{tasksvariable[i][2]}'
                                            planned = True
                                timesplanned = timesplanned + 1
                                if quarter + timesplanned == 52:
                                    break

                            if not planned:
                                nextemployee = nextemployee + 1
                                if nextemployee == len(workers):
                                    nextemployee = 0
                            if planned and timesplanned == (int(tasksvariable[i][5]) / 15):
                                workers[nextemployee][1] = workers[nextemployee][1] + 1
                                nextemployee = nextemployee + 1
                                if nextemployee == len(workers):
                                    nextemployee = 0
                                break
                        else:
                            nextemployee = nextemployee + 1
                            if nextemployee == len(workers):
                                nextemployee = 0


def priorities(person):
    workers = []
    for i, p in enumerate(person):
        if p[0].get():
            workers.append(i)
    random.shuffle(workers)
    nextworker = 0
    for prioritynumber in range(1, len(workers) + 1):
        prioritysetnumber = 0
        for i in range(52):
            for j in range(len(workers)):
                if prioritysetnumber == 4:
                    prioritysetnumber = 0
                    nextworker = nextworker + 1
                    if nextworker == len(workers):
                        nextworker = 0
                if person[workers[nextworker]][4][i][1] == 0 and not \
                        person[workers[nextworker]][4][i][0]['text'] and prioritysetnumber < 4:
                    person[workers[nextworker]][4][i][0]['text'] = prioritynumber
                    prioritysetnumber = prioritysetnumber + 1
                    break
                else:
                    prioritysetnumber = 0
                    nextworker = nextworker + 1
                if nextworker == len(workers):
                    nextworker = 0