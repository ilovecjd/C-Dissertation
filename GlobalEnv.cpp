
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




