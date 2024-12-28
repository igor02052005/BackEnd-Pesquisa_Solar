#include <stdio.h>

int global = 10;

void print(){
    int local = 20;

    printf("Global: %d\n", global);
    printf("Local: %d\n", local);
}

int main() {
    int local = 30;
    printf("Global: %d\n", global);
    printf("Local: %d\n", local);
    
    print();

    global = 40;
    printf("Global: %d\n", global);
    printf("Local: %d\n", local);
    return 0;
}

