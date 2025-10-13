def mini(a, b):
    """Return the minimum of a and b.
    """
    return a if a < b else b

def mini_n(a_list):
    """Return the minimum of the elements in a_list.
    """
    if len(a_list) < 3:
        return mini(a_list[0], a_list[1])
    else:
        return mini(mini_n(a_list[:-1]), a_list[-1])

def maxi(a, b):
    """Return the minimum of a and b.
    """
    return b if a < b else a

def maxi_n(a_list):
    """Return the minimum of the elements in a_list.
    """
    if len(a_list) < 3:
        return maxi(a_list[0], a_list[1])
    else:
        return maxi(maxi_n(a_list[:-1]), a_list[-1])


li = [1, 2, 3, 4, 5]
print(f"{maxi_n(li)}")
