def dot(a,b):
    return [[sum(a[i][k]*b[k][j] for k in range(len(a[0]))) for j in range(len(b[0]))] for i in range(len(a))]


a = [
        [1, 1, -2, 1],
        [1, 2, 3, -4],
        [2, 1, -1, -1],
        [1, -1, 1, 2]
    ]
b = [1, 2, 1, 3]
N = len(a)

l = [[1 if i == j else 0 for j in range(N)] for i in range(N)]
s = [r + [e] for (r,e) in zip(a,b)]
for r in s:
    print(r)

for k in range(N-1):
    maxMag = 0
    maxRow = k
    for m in range(k,N):
        if abs(s[m][k]) > maxMag:
            maxMag = abs(s[m][k])
            maxRow = m
    s[k], s[maxRow] = s[maxRow], s[k]
    l[k], l[maxRow] = l[maxRow], l[k]
    for i in range(k+1, N):
        factor = s[i][k] / s[k][k]
        l[i][k] = factor
        for j in range(k, N+1):
            s[i][j] = s[i][j] - factor * s[k][j]


#print("\nnew s")
#for r in s:
#    print(r[:-1])

s_restricted = [r[:-1] for r in s]
print("\ns_restricted")
for r in s_restricted:
    print(r)

print("\nl")
for r in l:
    print(r)

print("\nl @ s_restricted")
for r in dot(l, s_restricted):
    print(r)
