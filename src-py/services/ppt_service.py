"""
ppt_service.py
PowerPoint COM Interop 서비스 (pywin32 기반)

이벤트/수명 관리:
  - win32com.client.GetActiveObject 로 실행 중인 PPT 인스턴스에 연결.
  - COM 객체는 사용 후 del 로 참조 해제 (GC 의존 최소화).
  - PPT가 실행 중이 아니면 PptNotRunningError 발생.

스레드 안전성:
  - COM STA(Single-Threaded Apartment) 모델 준수.
  - 반드시 메인 스레드(또는 CoInitialize 호출된 스레드)에서 호출할 것.
  - tkinter 이벤트 루프 스레드에서 직접 호출 가능.

단축어 치환 전략:
  - 슬라이드 전체 순회 → TextFrame 보유 Shape → Run 단위 텍스트 검색.
  - 뒤에서 앞으로 치환하여 인덱스 오프셋 문제 방지.
"""
import re
from typing import List, Optional, Tuple

from models import SpecRecord

# ppSelectionText, ppSelectionShapes 상수 (PowerPoint 타입 라이브러리 없이 사용)
_PP_SELECTION_TEXT = 3
_PP_SELECTION_SHAPES = 2
_SHORTCODE_RE = re.compile(r"![A-Za-z0-9_\-]+")


class PptNotRunningError(Exception):
    """PowerPoint가 실행 중이 아닐 때 발생."""


class PptService:
    """실행 중인 PowerPoint 인스턴스에 연결하여 텍스트 삽입/치환을 수행."""

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @staticmethod
    def is_running() -> bool:
        """PowerPoint가 현재 실행 중인지 확인."""
        try:
            import win32com.client  # type: ignore
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            del app
            return True
        except Exception:
            return False

    @staticmethod
    def insert_text(text: str) -> bool:
        """
        현재 선택 영역(텍스트 커서 또는 도형)에 text를 삽입.

        반환값: True=성공, False=실패(PPT 미실행 또는 선택 없음)
        """
        try:
            import win32com.client  # type: ignore
            app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            return False

        try:
            win = app.ActiveWindow
            if win is None:
                return False

            sel = win.Selection
            sel_type = sel.Type

            if sel_type == _PP_SELECTION_TEXT:
                sel.TextRange.Text = text
                return True

            if sel_type == _PP_SELECTION_SHAPES:
                shapes = sel.ShapeRange
                if shapes.Count > 0:
                    shape = shapes.Item(1)
                    if shape.HasTextFrame:
                        tf = shape.TextFrame.TextRange
                        tf.Text = tf.Text + text
                        return True

            # 선택 없음 → 현재 슬라이드에 새 텍스트박스 추가
            slide = win.View.Slide
            new_shape = slide.Shapes.AddTextbox(1, 100, 100, 400, 50)
            new_shape.TextFrame.TextRange.Text = text
            return True

        except Exception as exc:
            print(f"[EquipSpec][ERROR] insert_text failed: {exc}")
            return False
        finally:
            del app

    @staticmethod
    def replace_shortcodes(
        records_lookup: dict,  # Dict[str, SpecRecord]
        scope: str = "active_slide",
    ) -> Tuple[int, List[str]]:
        """
        열린 프레젠테이션에서 !단축어를 일괄 치환.

        Parameters
        ----------
        records_lookup : dict
            {short_code: SpecRecord} 형태의 조회 딕셔너리
        scope : str
            "active_slide" | "all_slides"

        반환값: (치환 건수, 미발견 단축어 목록)
        """
        try:
            import win32com.client  # type: ignore
            app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            raise PptNotRunningError("PowerPoint가 실행 중이지 않습니다.")

        replaced_count = 0
        not_found: List[str] = []

        try:
            prs = app.ActivePresentation
            if prs is None:
                return 0, []

            if scope == "active_slide":
                slides = [app.ActiveWindow.View.Slide]
            else:
                slides = list(prs.Slides)

            for slide in slides:
                for shape in slide.Shapes:
                    if not shape.HasTextFrame:
                        continue
                    tf = shape.TextFrame.TextRange
                    full_text: str = tf.Text
                    matches = list(_SHORTCODE_RE.finditer(full_text))
                    if not matches:
                        continue

                    # 뒤에서 앞으로 치환 (인덱스 오프셋 방지)
                    for m in reversed(matches):
                        code = m.group(0)
                        rec: Optional[SpecRecord] = records_lookup.get(code)
                        if rec is None:
                            if code not in not_found:
                                not_found.append(code)
                            continue
                        start = m.start() + 1  # COM 1-based index
                        length = m.end() - m.start()
                        tf.Characters(start, length).Text = rec.display_value
                        replaced_count += 1

        except Exception as exc:
            print(f"[EquipSpec][ERROR] replace_shortcodes failed: {exc}")
        finally:
            del app

        return replaced_count, not_found

    @staticmethod
    def get_active_presentation_name() -> str:
        """현재 열린 프레젠테이션 파일명 반환. 없으면 빈 문자열."""
        try:
            import win32com.client  # type: ignore
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            name: str = app.ActivePresentation.Name
            del app
            return name
        except Exception:
            return ""
