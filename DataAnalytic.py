import pandas as pd
from pathlib import Path

ruta_csv = Path.home() / "Downloads" / "resumen_ventas.csv"

df = pd.read_csv(ruta_csv)
#print(df.head())

#productos_entre_50y70 = df[(df['precio_unitario'] >= 50) & (df['precio_unitario'] <= 70)]
#print(productos_entre_50y70)

def producto_pedido(id_pedido):
    producto = df.loc[df["id_pedido"] == id_pedido, "producto"]
    if producto.empty:
        return "Pedido no encontrado"
    
    return producto.iloc[0]

def precio_producto_exacto(producto):
    precio = df.loc[df["producto"] == producto, "precio_unitario"].iloc[0]
    return precio


print(producto_pedido(100))

