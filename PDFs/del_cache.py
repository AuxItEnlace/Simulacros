import shutil
import win32com.client.gencache

# Ruta a la caché
gen_py_path = win32com.client.gencache.GetGeneratePath()

print(f"🧹 Eliminando caché COM: {gen_py_path}")
shutil.rmtree(gen_py_path, ignore_errors=True)
print("✅ Caché eliminada.")
