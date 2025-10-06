def are_compatible(fam):
    """Check whether matrices in fam are compatible for mul.
    """
    mat_dims = []
    for m in fam:
        mat_dims.append((len(m), len(m[0])))

    for i in range(len(mat_dims) - 1):
        if mat_dims[i][1] != mat_dims[i+1][0]:
            return False
    return True

fam = [
        [[1, 1, 1],
        [1, 1, 1],
        [1, 1, 1],
        [1, 1, 1]],

        [[1, 2, 3, 4, 5],
        [1, 2, 3, 4, 5],
        [1, 2, 3, 4, 5]],

        [[3],
        [4],
        [5],
        [6],
        [7]]
]

fam2 = [
        [[1,2]],
        [[3,5]]
]

print(are_compatible(fam))
print(are_compatible(fam2))
#for r in mat_dims:
#    print(r)
