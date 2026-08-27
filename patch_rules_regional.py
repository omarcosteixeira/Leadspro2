import re

with open('firestore.rules', 'r') as f:
    rules = f.read()

# 1. Update isRestrictedRole
old_restricted = r"return !\(role in \['Admin Master', 'Gestor Comercial', 'Gerente Comercial \(Comercial\)', 'Financeiro'\]\);"
new_restricted = "return !(role in ['Admin Master', 'Gestor Comercial', 'Gerente Comercial (Comercial)', 'Financeiro', 'Regional']);"
rules = re.sub(old_restricted, new_restricted, rules)

# 2. Update isValidUser
old_valid_user = r"\['Admin Master', 'Promotor', 'FDV', 'Sala de Matrícula', 'QG', 'Líder/FDV', 'SSA', 'Gestor Unidade', 'Gestor Comercial', 'Acadêmico', 'Gerente Comercial \(Comercial\)', 'FDV \(Comercial\)', 'Promotor/rua', 'Financeiro', 'Técnico'\]"
new_valid_user = "['Admin Master', 'Promotor', 'FDV', 'Sala de Matrícula', 'QG', 'Líder/FDV', 'SSA', 'Gestor Unidade', 'Gestor Comercial', 'Acadêmico', 'Gerente Comercial (Comercial)', 'FDV (Comercial)', 'Promotor/rua', 'Financeiro', 'Técnico', 'Regional']"
rules = re.sub(old_valid_user, new_valid_user, rules)

# 3. Add isRegional()
old_is_tecnico = r"function isTecnico\(\) \{ return hasAnyRole\(\['Técnico'\]\); \}"
new_is_tecnico = "function isTecnico() { return hasAnyRole(['Técnico']); }\n    function isRegional() { return hasAnyRole(['Regional']); }"
rules = re.sub(old_is_tecnico, new_is_tecnico, rules)

# 4. Update linksUteis
old_links = r"allow write: if isPrincipal\(\) \|\| isComercial\(\) \|\| isSSA\(\);"
new_links = "allow write: if isPrincipal() || isComercial() || isSSA() || isRegional();"
rules = re.sub(old_links, new_links, rules)

# 5. Update forms_config
old_forms = r"allow write: if isPrincipal\(\) \|\| isComercial\(\);"
new_forms = "allow write: if isPrincipal() || isComercial() || isRegional();"
# be careful as this could replace other things. Let's do it specifically for forms_config block
old_forms_block = r"match /artifacts/[^/]+/public/data/forms_config/\{id\} \{\n\s*allow read: if true;\n\s*allow write: if isPrincipal\(\) \|\| isComercial\(\);\n\s*\}"
def replace_forms(match):
    return match.group(0).replace("allow write: if isPrincipal() || isComercial();", "allow write: if isPrincipal() || isComercial() || isRegional();")
rules = re.sub(old_forms_block, replace_forms, rules)

# 6. Update funcionarios
old_funcionarios_block = r"match /artifacts/[^/]+/public/data/funcionarios/\{id\} \{\n\s*allow read: if true;\n\s*allow write: if isAuthenticated\(\) && \(isPrincipal\(\) \|\| isComercial\(\)\);\n\s*\}"
def replace_func(match):
    return match.group(0).replace("isPrincipal() || isComercial()", "isPrincipal() || isComercial() || isRegional()")
rules = re.sub(old_funcionarios_block, replace_func, rules)

with open('firestore.rules', 'w') as f:
    f.write(rules)

print("Rules patched for Regional!")
