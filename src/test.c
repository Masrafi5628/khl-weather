#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <math.h>

int main()
{
    char name1[100],name2[100],remark[50];
    gets(name1);
    gets(name2);
    gets(remark);
    printf("Length1 - %d, Length2 - %d", strlen(name1),strlen(name2));
    if(strcmp(name1,name2)==0){
        printf("Exact Match\n");
    }
    else {
        printf("Do not match\n");
    }
    char name3[]="Ovi";

    strcpy(name1,name3);
    strcat(name1," ");
    strcat(name1,name2);
    strcat(name1," - ");
    puts(name1);
    strcat(name1,remark);

    puts(name1);



    return 0;
}