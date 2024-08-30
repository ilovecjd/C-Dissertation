
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

