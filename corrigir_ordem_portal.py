from pathlib import Path

p = Path("app.py")
t = p.read_text(encoding="utf-8")
lines = t.splitlines()

call_line = "render_portal_inicial()"
def_found = any(line.strip().startswith("def render_portal_inicial") for line in lines)

if not def_found:
    print("ERRO: a função def render_portal_inicial não foi encontrada no app.py")
    raise SystemExit(1)

# Comenta a primeira chamada solta que estiver antes da definição
def_index = next(i for i, line in enumerate(lines) if line.strip().startswith("def render_portal_inicial"))
changed = False

for i, line in enumerate(lines):
    if i < def_index and line.strip() == call_line:
        lines[i] = "# " + call_line + "  # movido para o final do arquivo"
        changed = True
        break

# Garante a chamada no final
if not any(line.strip() == call_line for line in lines[def_index+1:]):
    if lines and lines[-1].strip():
        lines.append("")
    lines.append(call_line)
    changed = True

if changed:
    p.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("Correção aplicada: chamada render_portal_inicial() movida para depois da definição.")
else:
    print("Nenhuma alteração necessária.")
