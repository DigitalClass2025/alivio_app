from pydantic import BaseModel
from typing import List, Optional

class Producto(BaseModel):
    sku: str
    nombre: str
    descripcion: Optional[str]
    precio: float
    stock: int
    categoria: Optional[str]