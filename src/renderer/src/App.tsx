import { useEffect, useMemo, useState } from 'react'
import type { ReactElement } from 'react'
import type { AppEnvironment, OneDriveLocation } from '@shared/types'
import './styles.css'

type LoadState =
  | { status: 'loading' }
  | { status: 'ready'; environment: AppEnvironment }
  | { status: 'error'; message: string }

const platformLabels: Record<string, string> = {
  darwin: 'macOS',
  win32: 'Windows',
  linux: 'Linux'
}

const sourceLabels: Record<OneDriveLocation['source'], string> = {
  'cloud-storage': 'CloudStorage',
  environment: '환경 변수',
  home: '홈 폴더',
  'user-profile': '사용자 프로필'
}

export function App(): ReactElement {
  const [loadState, setLoadState] = useState<LoadState>({ status: 'loading' })
  const [openingPath, setOpeningPath] = useState<string | null>(null)

  async function refreshEnvironment(): Promise<void> {
    setLoadState({ status: 'loading' })

    try {
      const environment = await window.oneDriveManager.getEnvironment()
      setLoadState({ status: 'ready', environment })
    } catch (error) {
      setLoadState({
        status: 'error',
        message: error instanceof Error ? error.message : '환경 정보를 확인하지 못했습니다.'
      })
    }
  }

  useEffect(() => {
    void refreshEnvironment()
  }, [])

  const content = useMemo(() => {
    if (loadState.status === 'loading') {
      return <div className="empty-state">로컬 OneDrive 경로를 확인하는 중입니다...</div>
    }

    if (loadState.status === 'error') {
      return <div className="empty-state error">{loadState.message}</div>
    }

    return (
      <>
        <section className="summary-grid" aria-label="실행 환경 요약">
          <SummaryTile label="플랫폼" value={platformLabels[loadState.environment.platform.name] ?? loadState.environment.platform.name} />
          <SummaryTile label="아키텍처" value={loadState.environment.platform.arch} />
          <SummaryTile label="감지된 경로" value={String(loadState.environment.oneDriveLocations.filter((location) => location.exists).length)} />
        </section>

        <section className="panel" aria-labelledby="paths-title">
          <div className="panel-heading">
            <div>
              <h2 id="paths-title">OneDrive 경로</h2>
              <p>{loadState.environment.platform.homeDirectory}</p>
            </div>
            <button className="secondary-button" type="button" onClick={() => void refreshEnvironment()}>
              새로고침
            </button>
          </div>

          <div className="path-list">
            {loadState.environment.oneDriveLocations.map((location) => (
              <PathRow
                key={location.path}
                location={location}
                isOpening={openingPath === location.path}
                onReveal={async () => {
                  setOpeningPath(location.path)

                  try {
                    await window.oneDriveManager.revealPath(location.path)
                  } finally {
                    setOpeningPath(null)
                  }
                }}
              />
            ))}
          </div>
        </section>

        <section className="panel compact" aria-labelledby="build-title">
          <div className="panel-heading">
            <div>
              <h2 id="build-title">빌드 대상</h2>
              <p>macOS DMG/ZIP, Windows NSIS/ZIP, x64 및 arm64를 지원합니다.</p>
            </div>
          </div>
        </section>
      </>
    )
  }, [loadState, openingPath])

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Desktop</p>
          <h1>OneDrive 관리자</h1>
        </div>
        <div className="status-pill">Windows + macOS</div>
      </header>

      {content}
    </main>
  )
}

function SummaryTile({ label, value }: { label: string; value: string }): ReactElement {
  return (
    <article className="summary-tile">
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  )
}

function PathRow({
  location,
  isOpening,
  onReveal
}: {
  location: OneDriveLocation
  isOpening: boolean
  onReveal: () => Promise<void>
}): ReactElement {
  return (
    <article className="path-row">
      <div className="path-status" data-state={location.exists ? 'found' : 'missing'} />
      <div className="path-copy">
        <div className="path-title">
          <strong>{location.label}</strong>
          <span>{sourceLabels[location.source]}</span>
        </div>
        <code>{location.path}</code>
      </div>
      <button className="primary-button" type="button" disabled={!location.exists || isOpening} onClick={() => void onReveal()}>
        {isOpening ? '여는 중' : '열기'}
      </button>
    </article>
  )
}
