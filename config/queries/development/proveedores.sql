SELECT
    cuenta_prov,
    nombre_prov,
    COUNT(DISTINCT pedido_compra) AS total_pedidos
FROM dbo.LineasPedidoProveedores
GROUP BY cuenta_prov, nombre_prov
ORDER BY nombre_prov