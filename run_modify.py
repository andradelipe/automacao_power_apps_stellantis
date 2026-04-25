import os

with open("teste.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if line.startswith('NUMERO_TAREFAS_REPETIDAS=""'):
        new_lines.append('NUMERO_TAREFAS_REPETIDAS=1\n')
        continue
        
    if 62 <= i <= 111:
        if i == 62:
            new_lines.append('            numero_repeticoes = int(NUMERO_TAREFAS_REPETIDAS) if str(NUMERO_TAREFAS_REPETIDAS).strip() else 1\n')
            new_lines.append('            for r in range(numero_repeticoes):\n')
            new_lines.append('                print(f"\\n--- Iniciando tarefa {r+1} de {numero_repeticoes} ---")\n')
            new_lines.append('                if r > 0:\n')
            new_lines.append('                    await asyncio.sleep(3)\n')
        
        new_lines.append('    ' + line)
    else:
        new_lines.append(line)

with open("teste.py", "w", encoding="utf-8") as f:
    f.writelines(new_lines)
