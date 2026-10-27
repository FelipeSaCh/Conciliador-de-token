def ordenar_por_prefijos(items, jerarquia):
    """Ordena por la posición del primer prefijo coincidente en `jerarquia`; el resto al final, alfabético."""
    def clave(item):
        s = str(item).strip()
        prioridad = next((i for i, p in enumerate(jerarquia) if s.startswith(p)), 999)
        return prioridad, s
    return sorted(items, key=clave)
