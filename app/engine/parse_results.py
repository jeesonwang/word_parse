from dataclasses import dataclass
from typing import List, Dict, Union

@dataclass
class WordProcessingResults:
    colored_pdf_path: str
    origin_pdf_path: str
    data: Dict[str, Union[str, List[dict]]]
