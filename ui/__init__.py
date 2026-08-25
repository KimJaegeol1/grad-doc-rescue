# -*- coding: utf-8 -*-
"""
ui — 화면.  판정 로직은 한 줄도 없다.

    theme.py    색 · 글꼴 · 로그 기호          검수도구 gui/theme.py 그대로
    runner.py   스레드 + 큐 + 중지            검수도구 gui/runner.py 그대로
    widgets.py  DropZone · PathRow · LogView  검수도구 gui/widgets.py 에서 셋만
    app.py      화면 한 장                    ★ 새로 쓴다

검수도구는 화면이 두 벌(gui/ 관리자용 · simple/ 사람용)이고 단계가 넷이라
이어달리기 셸(chain.py)이 필요했다.  이 도구는 **화면 한 장에 걸음 하나**다.
그래서 OptionBar · MatchTable · StepPanel 을 안 가져왔다.
"""
