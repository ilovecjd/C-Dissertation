
#include "stdafx.h"
#include "GlobalEnv.h"


GlobalEnv::GlobalEnv()
{
}


GlobalEnv::~GlobalEnv()
{
}

bool GlobalEnv::Init() {

	if (bIsInit ==0) {	// 초기화는 한번만 한다.
		return false;
	}
	else {
		
		// 현재 디렉토리를 저장할 버퍼
			char szPath[MAX_PATH];

		// 현재 디렉토리를 가져옵니다.
		GetCurrentDirectoryA(MAX_PATH, szPath);

		// 파일 이름을 결합하여 전체 경로 생성
		std::string strFilePath = std::string(szPath) + __STR_DATA_FILE;


		return true;
	}
	
}


//// Function to dynamically allocate a 2D array and initialize it to 0
//int** allocAndInit2DArray(int rows, int cols) {
//	// Step 1: Allocate memory for an array of int pointers (each pointer represents a row)
//	int** array = new int*[rows];
//
//	// Step 2: Allocate memory for each row and initialize to 0
//	for (int i = 0; i < rows; ++i) {
//		array[i] = new int[cols];
//		// Initialize each element in the row to 0
//		for (int j = 0; j < cols; ++j) {
//			array[i][j] = 0;
//		}
//	}
//
//	return array;
//}
//
//
//
//// Function to deallocate a dynamically allocated 2D array
//int** del2DArray(int** array, int rows) {
//	if (array != nullptr) {
//		// Deallocate memory for each row
//		for (int i = 0; i < rows; ++i) {
//			delete[] array[i];
//		}
//
//		// Deallocate the array of pointers
//		delete[] array;
//	}
//
//	// Return nullptr to assign to the pointer
//	return nullptr;
//}


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

