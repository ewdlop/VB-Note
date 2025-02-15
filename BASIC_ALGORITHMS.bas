REM Bubble Sort Algorithm
SUB BubbleSort(arr())
    DIM i AS INTEGER, j AS INTEGER, temp AS INTEGER
    FOR i = LBOUND(arr) TO UBOUND(arr) - 1
        FOR j = LBOUND(arr) TO UBOUND(arr) - i - 1
            IF arr(j) > arr(j + 1) THEN
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            END IF
        NEXT j
    NEXT i
END SUB

REM Binary Search Algorithm
FUNCTION BinarySearch(arr(), target AS INTEGER) AS INTEGER
    DIM low AS INTEGER, high AS INTEGER, mid AS INTEGER
    low = LBOUND(arr)
    high = UBOUND(arr)
    
    WHILE low <= high
        mid = (low + high) \ 2
        IF arr(mid) = target THEN
            BinarySearch = mid
            EXIT FUNCTION
        ELSEIF arr(mid) < target THEN
            low = mid + 1
        ELSE
            high = mid - 1
        END IF
    WEND
    BinarySearch = -1
END FUNCTION

REM Factorial Calculation
FUNCTION Factorial(n AS INTEGER) AS LONG
    IF n <= 1 THEN
        Factorial = 1
    ELSE
        Factorial = n * Factorial(n - 1)
    END IF
END FUNCTION

REM Greatest Common Divisor (GCD)
FUNCTION GCD(a AS INTEGER, b AS INTEGER) AS INTEGER
    DIM temp AS INTEGER
    WHILE b <> 0
        temp = b
        b = a MOD b
        a = temp
    WEND
    GCD = a
END FUNCTION

REM Linear Search Algorithm
FUNCTION LinearSearch(arr(), target AS INTEGER) AS INTEGER
    DIM i AS INTEGER
    FOR i = LBOUND(arr) TO UBOUND(arr)
        IF arr(i) = target THEN
            LinearSearch = i
            EXIT FUNCTION
        END IF
    NEXT i
    LinearSearch = -1
END FUNCTION

REM Fibonacci Series Generator
SUB Fibonacci(n AS INTEGER)
    DIM a AS LONG, b AS LONG, c AS LONG, i AS INTEGER
    a = 0
    b = 1
    PRINT a;
    PRINT b;
    FOR i = 3 TO n
        c = a + b
        PRINT c;
        a = b
        b = c
    NEXT i
END SUB

REM Prime Number Check
FUNCTION IsPrime(n AS INTEGER) AS INTEGER
    DIM i AS INTEGER
    IF n <= 1 THEN
        IsPrime = 0
        EXIT FUNCTION
    END IF
    IF n = 2 THEN
        IsPrime = 1
        EXIT FUNCTION
    END IF
    FOR i = 2 TO SQR(n)
        IF n MOD i = 0 THEN
            IsPrime = 0
            EXIT FUNCTION
        END IF
    NEXT i
    IsPrime = 1
END FUNCTION

REM String Reverse
FUNCTION ReverseString$(text AS STRING) AS STRING
    DIM i AS INTEGER, result AS STRING
    result = ""
    FOR i = LEN(text) TO 1 STEP -1
        result = result + MID$(text, i, 1)
    NEXT i
    ReverseString$ = result
END FUNCTION

REM Selection Sort
SUB SelectionSort(arr())
    DIM i AS INTEGER, j AS INTEGER, minIdx AS INTEGER, temp AS INTEGER
    FOR i = LBOUND(arr) TO UBOUND(arr) - 1
        minIdx = i
        FOR j = i + 1 TO UBOUND(arr)
            IF arr(j) < arr(minIdx) THEN
                minIdx = j
            END IF
        NEXT j
        IF minIdx <> i THEN
            temp = arr(i)
            arr(i) = arr(minIdx)
            arr(minIdx) = temp
        END IF
    NEXT i
END SUB

REM Example Usage Program
CLS
DIM numbers(9) AS INTEGER
DIM i AS INTEGER

REM Initialize array
FOR i = 0 TO 9
    numbers(i) = RND * 100
NEXT i

REM Print original array
PRINT "Original array:"
FOR i = 0 TO 9
    PRINT numbers(i);
NEXT i
PRINT

REM Sort array
CALL BubbleSort(numbers())

REM Print sorted array
PRINT "Sorted array:"
FOR i = 0 TO 9
    PRINT numbers(i);
NEXT i
PRINT

REM Test other functions
PRINT "Factorial of 5:", Factorial(5)
PRINT "GCD of 48 and 18:", GCD(48, 18)
PRINT "First 10 Fibonacci numbers:"
CALL Fibonacci(10)
PRINT
PRINT "Is 17 prime?", IsPrime(17)
PRINT "Reverse of 'BASIC':", ReverseString$("BASIC")
