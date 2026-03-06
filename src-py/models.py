"""
models.py
설비 사양 레코드 데이터 모델
"""
from dataclasses import dataclass, field


@dataclass
class SpecRecord:
    equip_id: str = ""
    short_code: str = ""
    spec_name: str = ""
    spec_value: str = ""
    unit: str = ""
    revision: int = 0

    @property
    def display_value(self) -> str:
        """삽입용 표시 문자열 (값 + 단위)"""
        if self.unit.strip():
            return f"{self.spec_value} {self.unit}"
        return self.spec_value

    def __str__(self) -> str:
        return (
            f"[{self.equip_id}] {self.short_code} | "
            f"{self.spec_name}: {self.display_value} (rev.{self.revision})"
        )
