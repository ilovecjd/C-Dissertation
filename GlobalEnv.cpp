
#include "stdafx.h"
#include "GlobalEnv.h"




int PoissonRandom(double lambda) {
	int k = 0;
	double p = 1.0;
	double L = exp(-lambda);  // L = e^(-lambda)

	do {
		k++;
		p *= static_cast<double>(rand()) / (RAND_MAX + 1.0);
	} while (p > L);

	return k - 1;
}

extern LPOLESTR gSheetNames[WS_NUM_SHEET_COUNT];// = { L"parameters", L"dashboard", L"project", L"activity_struct", L"debuginfo" };




												// 2차원 배열 동적 할당 함수
int** Newallocate2DArray(int rows, int cols) {
	int** array = new int*[rows]; // 행에 대한 포인터 배열을 동적 할당
	for (int i = 0; i < rows; i++) {
		array[i] = new int[cols]; // 각 행을 위한 int 배열을 동적 할당
	}
	return array;
}

// 2차원 배열 초기화 함수
void Newinitialize2DArray(int** array, int rows, int cols, int value) {
	for (int i = 0; i < rows; i++) {
		for (int j = 0; j < cols; j++) {
			array[i][j] = value; // 모든 요소를 value로 초기화
		}
	}
}

// 2차원 배열 동적 해제 함수
void Newdeallocate2DArray(int** array, int rows) {
	for (int i = 0; i < rows; i++) {
		delete[] array[i]; // 각 행에 대한 동적 메모리 해제
	}
	delete[] array; // 행 포인터 배열에 대한 동적 메모리 해제
}

// 2차원 배열 복사 함수
void Newcopy2DArray(int** source, int** destination, int rows, int cols) {
	for (int i = 0; i < rows; i++) {
		std::memcpy(destination[i], source[i], cols * sizeof(int));
	}
}


