import copy
from fractions import Fraction as fr


def gauss(u, b):
    a = copy.deepcopy(u)
    n = len(a)
    l = [[1 if i == j else 0 for j in range(n)] for i in range(n)]
    p = list(range(n))

    #Alg
    for k in range(n-1):
        print("---------------------------------------------------")
        print(f"{k=}")
        print("\nl")
        for r in l:
            print(r)
        print("\na")
        for r in a:
            print(r)

        kcolumn = [[ii, abs(a[ii][k])] for ii in range(k,n)]
        print(f"{kcolumn=}")
        ind_max = argmax(kcolumn)
        print(f"{ind_max=}")
        p = swap(p, ind_max, k)
        # Swap rows
        a[ind_max], a[k] = a[k], a[ind_max]
        b[ind_max], b[k] = b[k], b[ind_max]
        if k > 0:
            for x in range(k):
                l[k][x], l[ind_max][x] = l[ind_max][x], l[k][x]

        if k != ind_max:
            print("\nSwapped")
            print("\nl")
            for r in l:
                print(r)
            print("\na")
            for r in a:
                print(r)

        for i in range(k+1,n):
            # i)
            l[i][k] = a[i][k] / a[k][k]
            # ii)
            b[i] = b[i] - l[i][k] * b[k]
            # iii)
            for j in range(k+0,n):
                a[i][j] = a[i][j] - l[i][k] * a[k][j]

        print("\nAfter elimination")
        print(f"{k=}")
        print("\nl")
        for r in l:
            print(r)
        print("\na")
        for r in a:
            print(r)
        
    return l, a, b, p


def dot(a,b):
    return [[sum(a[i][k]*b[k][j] for k in range(len(a[0]))) for j in range(len(b[0]))] for i in range(len(a))]

def argmax(arr):
    biggest = max(arr[i][1] for i in range(len(arr)))
    result = arr[0][0]
    for i in range(0,len(arr)):
        if arr[i][1] == biggest:
            return arr[i][0]
    return result

def swap(p,i,j):
    p[i], p[j] = p[j], p[i]
    return p

def pinv(p):
    """Return the inverse of p."""
    result = [0 for i in range(len(p))]
    for i in range(len(p)):
        result[p[i]] = i
    return result

a = [
        [1, 1, -2, 1],
        [1, 2, 3, -4],
        [2, 1, -1, -1],
        [1, -1, 1, 2]
    ]
b = [1, 2, 1, 3]

#l, u = basic_gauss(a,b)
l, u, c, p = gauss(a,b)

print("\nl")
for r in l:
    print(r)

print("\nu")
for r in u:
    print(r)

print("\nl @ u")
for r in dot(l,u):
    print(r)

print("\np")
print(p)

print("\n p^t @ l @ u")
for i in range(len(l)):
    print(dot(l,u)[pinv(p)[i]])

aa = [
        [1, 0, 0],
        [3, -1, 0],
        [4, -1, -2]
    ]
bb = [
        [1, 1, 3],
        [0, 1, 5],
        [0, 0, 1]
    ]

#print("\nl @ u")
#for r in dot(l,u):
#    print(r)


#flag = False
#if flag:
#    u = [
#            [fr(1), fr(1), fr(-2), fr(1)],
#            [fr(1), fr(2), fr(3), fr(-4)],
#            [fr(2), fr(1), fr(-1), fr(-1)],
#            [fr(1), fr(-1), fr(1), fr(2)]
#        ]
#else:
#    u = [
#            [1, 1, -2, 1],
#            [1, 2, 3, -4],
#            [2, 1, -1, -1],
#            [1, -1, 1, 2]
#        ]
#
#
#u[0], u[2] = u[2], u[0]
##l[0], l[2] = l[2], l[0]
#k = 0
#if flag:
#    l = [[fr(1) if i == j else fr(0) for j in range(4)] for i in range(4)]
#else:
#    l = [[1 if i == j else 0 for j in range(4)] for i in range(4)]
#
#for i in range(k+1,4):
#    l[i][k] = fr(u[i][k] / u[k][k]) if flag else u[i][k] / u[k][k]
#    for j in range(0,4):
#        u[i][j] = u[i][j] - l[i][k] * u[k][j]
#print(f"\nu^({k})")
#for r in u:
#    print(r)
#
#k = 1
#for i in range(k+1,4):
#    l[i][k] = fr(u[i][k] / u[k][k]) if flag else u[i][k] / u[k][k]
#    for j in range(0,4):
#        u[i][j] = u[i][j] - l[i][k] * u[k][j]
#print(f"\nu^({k})")
#for r in u:
#    print(r)
#
#k = 2
#u[2], u[3] = u[3], u[2]
##l[2], l[3] = l[3], l[2]
#print(f"\nu^({k-1}) swapped")
#for r in u:
#    print(r)
#
#for i in range(k+1,4):
#    l[i][k] = fr(u[i][k] / u[k][k]) if flag else u[i][k] / u[k][k]
#    print(f"l({i},{k}) {l[i][k]}")
#    for j in range(0,4):
#        u[i][j] = u[i][j] - l[i][k] * u[k][j]
#print(f"\nu^({k})")
#for r in u:
#    print(r)
#
#
#print("\nl")
#for r in l:
#    print(r)
#
#print(f"\nl @ u^({k})")
#for r in dot(l,u):
#    print(r)
#
#del a, b, l, u, c
#a = [
#        [2, 1, -1, -1],
#        [1, 2, 3, -4],
#        [1, 1, -2, 1],
#        [1, -1, 1, 2]
#    ]
#b = [1, 2, 1, 3]

#l, u = basic_gauss(a,b)
#l, u  = basic_gauss(a,b)
#print("\nBasic")
#for r in l:
#    print(r)


def basic_gauss(u, b):
    n = len(a)
    l = [[1 if j==i else 0 for j in range(n)] for i in range(n)]
    a = copy.deepcopy(u)

    for k in range(n-1):
#        print(f"{k=}")
        for i in range(k+1,n):
            l[i][k] = a[i][k] / a[k][k]
            b[i] = b[i] - l[i][k] * b[k]
#            temp = copy.deepcopy(u)
            for j in range(k+0, n):
#                temp[i][j] = u[i][j] - l[i][k] * u[k][j]
                a[i][j] = a[i][j] - l[i][k] * a[k][j]
#            u = copy.deepcopy(temp)
    return l, a
