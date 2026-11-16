# R을 활용한 시험 문항 분석
# 2026학년도 1학기 중간고사 확률과 통계 선택형 21문항 (응시 252명)
# 광영고등학교 계기철
# 실행 환경: Posit Cloud (https://posit.cloud/) 또는 RStudio

# 0. 패키지 설치 (처음 한 번만) -----------------------------------------
install.packages(c("readxl", "ShinyItemAnalysis", "ltm"))

# 1. 데이터 불러오기 ------------------------------------------------------
# 정답 1, 오답 0 / 행 = 학생, 열 = 문항(1번~21번)
data1 <- read.csv("responses.csv", check.names = FALSE, fileEncoding = "UTF-8-BOM")
# 나이스 정오표를 정리한 엑셀을 쓸 때:
# library(readxl)
# data1 <- read_excel("data1.xlsx", sheet = "DATA")

dim(data1)               # 252 21
summary(rowSums(data1))  # 총점 요약

# 2. 고전검사이론(CTT): 난이도 · 변별도 ----------------------------------
library(ShinyItemAnalysis)

colMeans(data1)                        # 난이도(정답률) P = R / N
DDplot(data1)                          # 난이도 + 변별도(ULI, 상위 1/3 vs 하위 1/3)
DDplot(data1, k = 5, l = 4, u = 5)     # 5개 성적대 중 상위권(4) vs 최상위권(5)
ItemAnalysis(data1)                    # 문항별 지표 표

# 상·하위 27% 변별도 직접 계산
total <- rowSums(data1)
n27   <- round(nrow(data1) * 0.27)
ord   <- order(total)
low   <- data1[head(ord, n27), ]
high  <- data1[tail(ord, n27), ]
round(colMeans(high) - colMeans(low), 3)

# 3. 문항반응이론(IRT) ---------------------------------------------------
library(ltm)

fit2 <- ltm(data1 ~ z1)       # 2모수 모형
coef(fit2)                    # Dffclt = 난이도 b, Dscrmn = 변별도 a
plot(fit2, type = "ICC")      # 문항특성곡선(ICC)
plot(fit2, type = "ICC", items = c(5, 9, 16))   # 변별도 낮은 문항만

fit3 <- tpm(data1)            # 3모수 모형 (추측도 c 포함)
coef(fit3)

anova(fit2, fit3)             # 2모수와 3모수 모형 비교
