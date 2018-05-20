#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <windows.h>

//ループ用変数
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

//bubblesort用カウント
int SortCount;
//配列サイズ入力用（用途多岐）
int ArraySize;
//配列宣言
int *Array;
//ランダム数値生成の際の桁数用
int digits;

//プロトタイプ宣言
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

//入力された数のint型配列を作成する
int ArrayMaker(){
    printf("確保する配列のサイズを指定してください。:");
    scanf("%d",&ArraySize);
    Array = (int *)malloc(sizeof(int) * ArraySize);
}

//int型配列にrandomな数値を入力する
int RandNumGiveToArray(){
    printf("生成するrandom数値の最大桁数を入力してください :");
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