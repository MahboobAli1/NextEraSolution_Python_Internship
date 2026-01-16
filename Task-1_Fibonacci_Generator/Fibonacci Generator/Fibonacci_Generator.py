def fibo(start, end):
    if start <= 0 or end <= 0:
        print("Cannot generate Fibonacci for non-positive values")
        return
    elif end<start:
        print("start value can not be grater than end ")

    a, b = 0, 1

    while a <= end:
        if a >= start:
            print(a, end=" ")
        a, b = b, a + b
a=int(input("Enter starting number:"))
b= int(input("Enter ending number:"))
fibo(a,b)
