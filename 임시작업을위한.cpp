/*

RND_HR_H -> 변수로 바꾸자

1)
int CProject::CalculateHRAndProfit() {
	int high = 0, mid = 0, low = 0;

	for (int i = 0; i < numActivities; ++i) {
		==> 각 스킬은 한명을 제한됨. 각각 여러명이 투입되게 수정하자.

2)
int CProject::CalculateLaborCost(const std::string& grade) {
case 'H':
		directLaborCost = 50; ==> 동적으로 	바뀔수 있게 변경하자.
}		

3)
선금, 중도금, 잔금은 double 로 계산하고 표시는 정수로 하고 차액은 보정하자
추후 기대수익으로 변환해서 결정하는 순간의 가치로 평가하자.