# AI Agent Platform

사내 업무 자동화를 위한 통합 Agent Runtime

# 프로젝트 배경

## 문제 정의

반복 업무가 팀의 시간을 잠식하고 있다.

- 접수·분류·전달이 모두 수동
- 시스템마다 다른 인터페이스
- 처리 이력이 흩어져 감사 불가

## 핵심 지표

| 지표 | 현재 | 목표 |
|---|---:|---:|
| 가용성 | 99.9% | 99.99% |
| 비용 절감 | 12% | 43% |
| 배포 속도 | 1.0x | 2.4x |

# 아키텍처

## 핵심 가치

### Automation

반복 작업을 Agent가 대신한다.

### Integration

MCP로 사내 시스템을 연결한다.

### Intelligence

LLM이 맥락을 읽고 판단한다.

## 주요 성과

- 99.99% Availability
- 43% Cost Reduction
- 2.4x Deployment Speed

## IRSA vs Pod Identity

### IRSA

- 표준 방식, 넓은 생태계 지원
- 어노테이션 기반 설정

### Pod Identity

- 간단한 설정, 신규 권장
- 클러스터 단위 관리

## 구성 요소

### Control Plane

정책, 권한, 감사 로그를 관리한다.

### Data Plane

Agent 실행과 도구 호출을 담당한다.

```ts
const agent = new Agent({ tools: mcpTools });
await agent.run(task);
```

# 도입 계획

## 현장의 목소리

> 도입 후 반복 업무가 사라지고, 팀은 판단이 필요한 일에 집중하게 됐다.

— 운영팀 리드

## 도입 절차

1. 접수 자동 분류
2. 초안 자동 생성
3. 담당자 검토
4. 자동 발송

## 다음 단계

- 파일럿 팀 확대
- 보안 검토 완료
- 전사 배포

## 감사합니다

문의: platform@example.com
