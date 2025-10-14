import sys

def exp_by_sq(x, n):
    if n < 0:
        x = 1 / x
        n = -n
    if n == 0:
        return 1
    y = 1
    while n > 1:
        if n % 2 == 1:
            y = x * y
            n -= 1
        x = x * x
        n = n / 2
    return x * y

x = int(sys.argv[1])
n = int(sys.argv[2])
print(f"{x} ^ {n} = {exp_by_sq(x,n)}")
