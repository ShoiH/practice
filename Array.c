#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <windows.h>

//���[�v�p�ϐ�
int i = 0;
int j = 0;
int k = 0;
int l = 0;
int m = 0;
int n = 0;
int o = 0;
int p = 0;
int q = 0;
int r = 0;
int s = 0;
int t = 0;

//bubblesort�p�J�E���g
int SortCount;
//�z��T�C�Y���͗p�i�p�r����j
int ArraySize;
//�z��錾
int *Array;
//�����_�����l�����̍ۂ̌����p
int digits;

//�v���g�^�C�v�錾
int ArrayMaker();
int RandNumGiveToArray();
int PrintArray();
int BubbleSort();

int main(void){
    ArrayMaker();
    RandNumGiveToArray();
    PrintArray();
    BubbleSort();
    PrintArray();
    free(Array);
    return 0;
}

//���͂��ꂽ����int�^�z����쐬����
int ArrayMaker(){
    printf("�m�ۂ���z��̃T�C�Y���w�肵�Ă��������B:");
    scanf("%d",&ArraySize);
    Array = (int *)malloc(sizeof(int) * ArraySize);
}

//int�^�z���random�Ȑ��l����͂���
int RandNumGiveToArray(){
    printf("��������random���l�̍ő包������͂��Ă������� :");
    scanf("%d",&digits);

    srand((int)time(NULL));

    int pow = 1;

    for(l=1;l<=digits;l++){
        pow = pow * 10;
    }

    for(i=0;i<ArraySize + 1;i++){
        for(k=0;k<6;k++){
            rand();
        }
        Array[i] = rand() % pow;
    }
}

int PrintArray(){
    printf("{");
    for(j=0;j<ArraySize;j++){
        printf("%d,",Array[j]);
    }
    printf("}\n");
}

int BubbleSort(){
    printf("BubbleSort\n");
    SortCount = 1;
    if(SortCount != 0){
        for(m=0;m<ArraySize;m++){
            SortCount = 0;
            for(n=ArraySize-1;n>m;n--){
                if(Array[n]<Array[m]){
                o = Array[n];
                Array[n]=Array[m];
                Array[m]=o;
                SortCount++;
                }
            }
        }
    }
}

//p,q,r,s,t

// int Pivot(){

// }

// int QuickSort(){
// }