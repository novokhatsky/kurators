import math

str1 = 'я пришёл к себе домой'
str2 = 'я пришел домой к себе'


# возвращает отсортированные строки, первой идет с максимальной длиной
def max_str(str1, str2):

    if len(str1) > len(str2):
        return str1, str2

    return str2, str1


# возвращает длину максимальной строки
def max_len(str1, str2):

    if len(str1) > len(str2):
        return len(str1)

    return len(str2)


# вычисляет длину совпадений l
def len_matches(len_max):
    return math.floor(len_max / 2) - 1


def jaro(str1, str2):
    str1, str2 = max_str(str1, str2)
    l = len_matches(max_len(str1, str2))
    e = 0
    z = 0

    for i in range(len(str2)):

        # если точное совпадение, увеличиваем счетчик и к следующему символу
        if str2[i] == str1[i]:
            e += 1
            continue

        # попробуем найти совпадение на расстоянии, не дальше -l, +l
        for j in range(l * (-1), l + 1):

            if j == 0:
                continue

            try:
                if str1[i] == str2[i + j]:
                    z += 1
                    break
            except IndexError:
                continue

    m = e + z
    t = z / 2

    if m == 0:
        return 0

    return int(100 * ((m / len(str1)) + (m / len(str2)) + ((m - t)/m)) / 3)


def karlovskiy_distance(str1, str2):

    str1 = '\t\t' + str1 + '\t\t'
    str2 = '\t\t' + str2 + '\t\t'

    dist = -4

    for i in range(len(str1) - 2):
        if str1[i:i+3] not in str2:
            dist += 1

    for i in range(len(str2) - 2):
        if str2[i:i+3] not in str1:
            dist += 1

    return int((1 - dist/(len(str1) + len(str2) - 8)) * 100)


def lev(str1, str2):
    if len(str1) > len(str2):
        str1, str2 = str2, str1

    n = len(str1)
    m = len(str2)

    current_row = range(n + 1)

    for i in range(1, m + 1):
        previous_row, current_row = current_row, [i] + [0] * n

        for j in range(1, n + 1):
            add = previous_row[j] + 1
            delete = current_row[j - 1] + 1
            change = previous_row[j - 1]

            if str1[j - 1] != str2[i - 1]:
                change += 1

            current_row[j] = min(add, delete, change)
            pass

    return 100 - int(current_row[n] * 100 / len(str2))


print(str1)
print(str2)
print("джаро {0}".format(jaro(str1, str2)))
print("карловский {0}".format(karlovskiy_distance(str1, str2)))
print("Левенштейна {0}".format(lev(str1, str2)))