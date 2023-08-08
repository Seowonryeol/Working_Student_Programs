import os
import datetime


try:
    num=int(input("숫자를 입력하세요: "))
    result=10/num
    print(result)
except ZeroDivisionError:
    print("0으로 나눌 수 없음")
    print("0으로 나눌 수 없음")
finally:
    print("프로그램이 종료되었습니다")
