
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


/////////////////////////////////////////////////////////////////////
// 파일 처리 루틴들
bool OpenFile(const CString& filename, const TCHAR* mode, FILE** fp) {
	errno_t err = _wfopen_s(fp, filename, mode);
	if (err != 0 || *fp == nullptr) {
		perror("Failed to open file");
		return false;
	}
	return true;
}

void CloseFile(FILE** fp) {
	if (*fp != nullptr) {
		fclose(*fp);
		*fp = nullptr;
	}
}

ULONG WriteDataWithHeader(FILE* fp, int type, const void* data, size_t dataSize) {

	ULONG ulWritten = 0;
	ULONG ulTemp = 0;
	SAVE_TL tl = { type, static_cast<int>(dataSize) };
	ulTemp = fwrite(&tl, sizeof(tl), 1,fp);  // 먼저 데이터 타입 및 길이 정보를 쓴다
	ulWritten += ulTemp * sizeof(tl);

	ulTemp = fwrite(data, dataSize,1, fp);   // 실제 데이터 쓰기
	ulWritten += ulTemp * dataSize;

	return  ulWritten;
}

bool ReadDataWithHeader(FILE* fp, void* data, size_t expectedSize, int expectedType) {
	SAVE_TL tl;
	if (fread(&tl, 1, sizeof(tl), fp) != sizeof(tl)) {
		perror("Failed to read header");
		return false;
	}

	if (tl.type != expectedType || tl.length != expectedSize) {
		fprintf(stderr, "Data type or size mismatch\n");
		return false;
	}

	if (fread(data, 1, tl.length, fp) != tl.length) {
		perror("Failed to read data");
		return false;
	}

	return true;
}

