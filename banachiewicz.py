def lu(arr):
    row_count = len(arr)

    # L - zero, U - unit
    L_array = [[0 for j in range(row_count)]
               for i in range(row_count)]
    U_array = [[1 if i == j else 0 for j in range(row_count)]
               for i in range(row_count)]
    
    # 3.
    for i in range(row_count):
        L_array[i][0] = arr[i][0]
    # 4.
    for j in range(1,row_count):
        U_array[0][j] = arr[0][j] / L_array[0][0]

    # 5.
    for j in range(1,row_count):
        # a)
        for i in range(1,j): # Upper bound might be a problem
            U_array[i][j] = sum(L_array[i][k] * U_array[k][j]
                               for k in range(i))
            U_array[i][j] = (arr[i][j] - U_array[i][j]) / L_array[i][i]
        # b)
        for i in range(j, row_count):
            L_array[i][j] = sum(L_array[i][k] * U_array[k][j]
                               for k in range(j))
            L_array[i][j] = arr[i][j] - L_array[i][j]
    
    return (L_array, U_array)

#arr = [[1, 1, 3],
#       [3, 2, 4],
#       [4, 3, 5]]

n = 10
arr = [[0 for j in range(n)] for i in range(n)]
k = 1
for i in range(n):
    for j in range(n):
        arr[i][j] = k
        k += 1
print("\narr")
for r in arr:
    print(r)
l, u = lu(arr)
print("\nl")
for r in l:
    print(r)
print("\nu")
for r in u:
    print(r)
