# AI Agent Platform

사내 업무 자동화를 위한 통합 Agent Runtime

## 프로젝트 개요

AgentDure는 사내 시스템을 하나의 Agent Runtime으로 연결한다.

- MCP 기반 도구 연동
- LLM 기반 판단과 실행
- 감사 가능한 실행 이력

## 핵심 지표

| 지표 | 현재 | 목표 |
|---|---:|---:|
| 가용성 | 99.9% | 99.99% |
| 비용 절감 | 12% | 43% |
| 배포 속도 | 1.0x | 2.4x |

## 아키텍처

### Control Plane

정책, 권한, 감사 로그를 관리한다.

### Data Plane

Agent 실행과 도구 호출을 담당한다.

```ts
const agent = new Agent({ tools: mcpTools });
await agent.run(task);
```

## 도입 효과

> 반복 업무가 사라지고, 팀은 판단이 필요한 일에 집중한다.

1. 접수 자동 분류
2. 초안 자동 생성
3. 담당자 검토
4. 자동 발송

## 다음 단계

- 파일럿 팀 확대
- 보안 검토 완료
- 전사 배포
