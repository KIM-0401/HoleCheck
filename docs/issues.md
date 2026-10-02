# 겪은 문제 누적

(CLAUDE.md 형식으로 append · "왜 예측 못 했나" 필수)

## W04 · src 아래 VS 프로젝트가 두 개 생김 (HoleCheck · HoleCheckAddIn)
- 증상: src/HoleCheck/(17:06 생성)와 src/HoleCheckAddIn/(18:10 생성)이 둘 다 있다. 둘 다 Class1.vb만 들어 있고 sln도 각각 따로 있다. W04 기준은 src\HoleCheckAddIn 단일 프로젝트
- 원인: (추정 · 확인 필요) 처음에 프로젝트 이름을 HoleCheck로 만들었고, 이름을 바꾸지 않고 새로 만든 뒤 이전 폴더를 지우지 않음
- 처리: src/HoleCheck 폴더를 사람이 삭제함 (W04 마감 다음 날 확인). src/HoleCheckAddIn 하나만 남음
- 왜 예측 못 했나: VS 새 프로젝트 대화상자에서 프로젝트 이름이 곧 폴더명·AssemblyName·RootNamespace가 된다는 점을 생성 전에 확인하지 않았다. 과제 파일도 경로만 적고 프로젝트 이름은 명시하지 않았다

## W04 · 생성된 vbproj 설정이 CLAUDE.md 규칙과 다름
- 증상: HoleCheckAddIn.vbproj가 OptionStrict=Off이고 기본 Imports(System, Microsoft.VisualBasic, System.Linq 등 9개)가 켜져 있다. Inventor Interop 참조도 아직 없다
- 원인: VS 클래스 라이브러리 템플릿 기본값 그대로임. 프로젝트 속성 > 컴파일 / 참조 단계를 아직 진행하지 않음
- 처리: 미처리. W04 B 항목(Option Strict=On, Interop 참조 Embed=False/CopyLocal=False)과 CLAUDE.md(기본 Imports 끔)에 맞춰 VS 속성 창에서 사람이 바꿔야 함
- 왜 예측 못 했나: 템플릿 기본값이 Strict Off이고 Imports가 켜진 상태라는 걸 몰랐다. 프로젝트를 만든 시점에 CLAUDE.md 기술 스택 항목과 대조하는 체크 단계가 없었다
