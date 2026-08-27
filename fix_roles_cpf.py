import re

with open('src/App.tsx', 'r') as f:
    app = f.read()

# Fix 1: Add "Regional" to isComercial allowed roles
# The code is:
#                           if (isComercial) {
#                             return [
#                               "Admin Master",
#                               "Gerente Comercial (Comercial)",
#                               "FDV (Comercial)",
#                               "Promotor/rua",
#                               "Financeiro",
#                             ].includes(r);
#                           } else {

old_comercial = r'if \(isComercial\) \{\s*return \[\s*"Admin Master",\s*"Gerente Comercial \(Comercial\)",\s*"FDV \(Comercial\)",\s*"Promotor/rua",\s*"Financeiro",\s*\]\.includes\(r\);\s*\} else \{'
new_comercial = '''if (isComercial) {
                            return [
                              "Admin Master",
                              "Gerente Comercial (Comercial)",
                              "FDV (Comercial)",
                              "Promotor/rua",
                              "Financeiro",
                              "Regional"
                            ].includes(r);
                          } else {'''
app = re.sub(old_comercial, new_comercial, app)


# Fix 2: Add "Regional" to NOT isComercial allowed roles
#                           } else {
#                             return ![
#                               "Gerente Comercial (Comercial)",
#                               "FDV (Comercial)",
#                               "Promotor/rua",
#                             ].includes(r);
#                           }
# We don't need to change `else` because it returns things that are NOT in the array. 
# So Regional is already allowed in Principal server.

# Fix 3: Remove CPF duplicate block in FiesProuniView
cpf_block = r'const isDuplicate = data\.some\(\s*\(item\) => item\.cpf === cleanCpf && item\.id !== editingEntry\?\.id,\s*\);\s*if \(isDuplicate\) \{\s*onToast\("Este CPF já está cadastrado no FIES/Prouni\.", "error"\);\s*return;\s*\}'

app = re.sub(cpf_block, '', app)

with open('src/App.tsx', 'w') as f:
    f.write(app)

print("Fixed roles and CPF!")
