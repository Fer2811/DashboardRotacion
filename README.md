# Dashboard IQ Tech · VNext Repartición + Odoo

Fuente de verdad: dashboard con **Repartición + B2B mínimo 10%**.

Implementado:
- SKU duplicados resueltos usando stock real por `product_id` en la ubicación origen.
- Varios lotes del mismo producto NO se tratan como producto duplicado.
- Sugerencias: primero **Evaluar movimiento**, luego botón Odoo dentro de la revisión.
- Repartición restaurada completa y con borradores Odoo por destino/fuente.
- Backend preparado para una prueba controlada de Odoo→Full como **Entrega/outgoing**, desactivada por defecto.
- Todos los movimientos automáticos de esta fase quedan en **Borrador**.

## Roles propuestos para fase multiusuario
- Lectura: consulta solamente.
- Solicitante: consulta + crea solicitudes compartidas.
- Operador: revisa/aprueba solicitudes y crea borradores Odoo.
- Administrador: Operador + gestión de usuarios/roles/configuración/auditoría.

Las solicitudes compartidas deben pasar de `localStorage` a una base del servidor (SQLite es suficiente para esta etapa).

## Ejecutar
`python run_dashboard.py` y abrir `http://127.0.0.1:5001`.
