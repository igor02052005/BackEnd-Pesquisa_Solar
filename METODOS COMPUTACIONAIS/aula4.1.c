#include <stdio.h>

void add(int a, int b){
    int res = a + b;
    printf("Sum: %d\n", res);
}

int main(){
    int num1 = 5, num2 = 10;
    add(num1, num2);

    for (int i = 0; i <= 5; i++)
    {
        printf("%d * %d = %d\n", num1, i, num1 * i);
    }
    
    return 0;
}
