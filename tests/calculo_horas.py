from datetime import datetime, timedelta

def calcular_horas(periodos):
    total_horas = 0
    periodos = periodos.split()  # Divide los diferentes bloques de tiempo

    for periodo in periodos:
        inicio, fin = periodo.split('-')  # Separa las horas de inicio y fin
        fmt = '%H:%M:%S'  # Formato de las horas
        t_inicio = datetime.strptime(inicio, fmt)
        t_fin = datetime.strptime(fin, fmt)

        # Si la hora de fin es menor que la de inicio, sumamos 1 día (24 horas)
        if t_fin < t_inicio:
            t_fin += timedelta(days=1)

        # Calcula la diferencia en horas y añade al total
        horas = (t_fin - t_inicio).total_seconds() / 3600
        total_horas += horas

    return total_horas

# Ejemplo de uso
print(calcular_horas('11:00:00-16:00:00 20:00:00-00:00:00'))  # Resultado esperado: 9 horas
print(calcular_horas('06:00:00-11:00:00 16:00:00-02:00:00'))  # Resultado esperado: 9 horas
