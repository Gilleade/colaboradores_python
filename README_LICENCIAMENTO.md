# README de Licenciamento — Projeto Colaboradores

## 1. Finalidade deste documento

Este documento registra o fluxo oficial de licenciamento do projeto **Colaboradores**, incluindo emissão, instalação, validação, renovação e boas práticas operacionais. O objetivo é manter um procedimento padronizado, seguro e rastreável, evitando perda de controle sobre versões distribuídas do sistema.

Este material é de uso **interno do desenvolvedor** e deve permanecer no projeto como referência operacional.

---

## 2. Visão geral do modelo de licenciamento

O sistema utiliza um modelo de licença **offline**, baseado em:

- arquivo de licença externo;
- assinatura criptográfica da licença;
- validação na abertura do programa;
- bloqueio do sistema quando a licença estiver ausente, inválida, expirada ou incompatível com a máquina;
- solicitação técnica de ativação por meio de `activation_request.json`.

A proteção real da licença não depende de esconder o arquivo, mas da **assinatura criptográfica**. O usuário pode visualizar o conteúdo da licença, porém não consegue alterar dados relevantes sem invalidar a assinatura.

---

## 3. Componentes do licenciamento

### 3.1. Arquivos privados do desenvolvedor

Esses arquivos devem permanecer sob controle exclusivo do desenvolvedor:

- `tools_private/generate_keys.py`
- `tools_private/emitir_licenca.py`
- `tools_private/keys/private_key.pem`

**Regra crítica:** `private_key.pem` nunca deve ser enviada ao cliente, nunca deve acompanhar o build e nunca deve ser publicada em repositórios remotos.

### 3.2. Arquivos usados pelo programa

Esses arquivos participam da validação da licença:

- `license_validator.py`
- `public_key.pem`

### 3.3. Arquivos distribuídos ou gerados no lado do cliente

- `license.json` (ou outro arquivo de licença gerado pelo emissor)
- `activation_request.json`

---

## 4. Tipos de licença suportados

Atualmente o sistema suporta dois tipos de licença.

### 4.1. Licença temporária

Valor interno:

```json
"license_type": "temporary"
```

Características:

- exige data de expiração;
- bloqueia automaticamente após o vencimento;
- indicada para testes, uso provisório, regularização ou renovação periódica.

Campo obrigatório:

- `expires_at`

### 4.2. Licença vitalícia

Valor interno:

```json
"license_type": "perpetual"
```

Características:

- não expira;
- indicada para liberação definitiva;
- pode ou não ser vinculada a uma máquina específica.

Campo esperado:

- `expires_at: null`

---

## 5. Escopos de licenciamento suportados

O emissor exige que o escopo seja definido explicitamente pelo desenvolvedor. Os escopos suportados são:

- `person` → licença para uma pessoa específica;
- `sector` → licença para um setor específico;
- `company` → licença para a empresa como um todo.

Esses dados são sempre definidos pelo desenvolvedor no momento da emissão. O usuário final não escolhe escopo, validade, tipo de licença nem vínculo por máquina.

---

## 6. Vínculo com máquina (`machine_id`)

O sistema suporta licenças vinculadas a uma máquina específica por meio do campo `machine_id`.

Quando a licença contém esse campo preenchido, o programa compara o identificador gravado na licença com o identificador real da máquina em execução. Se os valores forem diferentes, a licença é recusada.

### 6.1. Recomendações práticas

Uso recomendado:

- **licença temporária**: normalmente sem vínculo por máquina;
- **licença vitalícia**: preferencialmente com vínculo por máquina.

Esse critério pode ser alterado conforme a estratégia comercial adotada.

---

## 7. Estrutura do `activation_request.json`

O arquivo `activation_request.json` foi definido para conter **apenas dados técnicos**, sem qualquer decisão comercial.

Exemplo esperado:

```json
{
  "app_id": "colaboradores",
  "app_name": "Separador de Listas e Provas",
  "machine_id": "...",
  "generated_at": "2026-03-26T08:40:47",
  "expected_license_path": "C:\\...\\license.json"
}
```

Finalidade:

- informar ao desenvolvedor a máquina de destino;
- permitir emissão de licença vinculada corretamente;
- registrar data/hora da solicitação e caminho esperado da licença.

O usuário final não deve preencher manualmente esse arquivo.

---

## 8. Geração das chaves

A geração das chaves é feita uma única vez, salvo necessidade de reinicialização completa do sistema de licenciamento.

### 8.1. Comando

No PowerShell, a partir da raiz do projeto:

```powershell
python tools_private/generate_keys.py
```

### 8.2. Resultado esperado

Arquivos gerados:

- `tools_private/keys/private_key.pem`
- `tools_private/keys/public_key.pem`

### 8.3. Regras de segurança

- `private_key.pem` é sigilosa e deve permanecer somente com o desenvolvedor;
- `public_key.pem` é a chave pública usada pelo programa para validar a assinatura;
- a perda da `private_key.pem` inviabiliza a emissão de novas licenças compatíveis com as já distribuídas.

---

## 9. Emissão de licenças

A emissão é realizada pelo script:

- `tools_private/emitir_licenca.py`

### 9.1. Campos definidos pelo desenvolvedor

Na emissão, o desenvolvedor decide:

- `license_id`
- `customer_type`
- `customer_name`
- `license_type`
- `expires_at` (quando temporária)
- uso ou não de `machine_id`
- observações (`notes`)

### 9.2. Exemplo — licença temporária sem vínculo por máquina

```powershell
python tools_private/emitir_licenca.py --license-id LIC-TESTE-0003 --customer-type sector --customer-name "Setor de Teste" --license-type temporary --expires-at 2026-04-02 --notes "Temporária sem vínculo de máquina" --output license_temp.json
```

### 9.3. Exemplo — licença vitalícia vinculada à máquina do pedido

```powershell
python tools_private/emitir_licenca.py --license-id LIC-PERP-0002 --customer-type person --customer-name "Nicolly" --license-type perpetual --notes "Permanente vinculada à máquina" --activation-request activation_request.json --bind-to-activation-machine --output license_permanent_bound.json
```

### 9.4. Exemplo — licença temporária vinculada à máquina do pedido

```powershell
python tools_private/emitir_licenca.py --license-id LIC-TEMP-0005 --customer-type company --customer-name "Empresa X" --license-type temporary --expires-at 2026-04-15 --notes "Temporária vinculada à máquina" --activation-request activation_request.json --bind-to-activation-machine --output license_temp_bound.json
```

### 9.5. Exemplo — licença com `machine_id` manual

```powershell
python tools_private/emitir_licenca.py --license-id LIC-MANUAL-0001 --customer-type person --customer-name "Cliente X" --license-type perpetual --machine-id HASH_AQUI --notes "Vitalícia manual" --output license_manual.json
```

### 9.6. Observações importantes

- o uso de `activation_request.json` é preferível ao preenchimento manual do `machine_id`;
- `customer_name` e `customer_type` devem sempre ser definidos explicitamente pelo desenvolvedor;
- o vínculo por máquina só ocorre quando for realmente solicitado na emissão.

---

## 10. Convenção recomendada para `license_id`

É recomendável manter um padrão consistente de identificação de licenças. Exemplos:

- `LIC-2026-0001`
- `LIC-TEMP-0007`
- `LIC-PERP-0002`
- `LIC-EMPRESA-0001`

Sugestão prática:

- usar prefixos que indiquem o tipo da licença;
- manter numeração sequencial;
- evitar reaproveitamento de identificadores.

---

## 11. Instalação da licença no lado do cliente

O fluxo implementado no programa é o seguinte:

1. o programa inicia;
2. se a licença estiver ausente, inválida, expirada ou incompatível, a interface principal não é aberta;
3. a janela de bloqueio exibe a mensagem de licenciamento e permite:
   - visualizar a `machine_id`;
   - gerar `activation_request.json`;
   - instalar uma nova licença;
4. o usuário clica em **Instalar licença**;
5. seleciona o arquivo de licença fornecido pelo desenvolvedor;
6. o programa copia o arquivo para o local esperado;
7. a assinatura é validada automaticamente;
8. se a licença for válida, o usuário recebe confirmação e pode abrir o programa novamente.

---

## 12. Validação da licença pelo programa

Atualmente o programa valida os seguintes pontos:

- existência do arquivo de licença;
- integridade do JSON;
- presença dos campos obrigatórios;
- assinatura criptográfica;
- data de expiração, quando aplicável;
- compatibilidade com a `machine_id`, quando aplicável;
- retrocesso suspeito da data/hora do sistema.

Se qualquer uma dessas verificações falhar, o uso do programa é bloqueado.

---

## 13. Proteção contra alteração manual da licença

O arquivo de licença é assinado. Portanto:

- o usuário pode abrir o arquivo;
- o usuário pode editar o arquivo;
- mas qualquer alteração relevante invalida a assinatura.

Na prática, isso significa que a proteção não depende de ocultar a licença, e sim da impossibilidade de gerar uma nova assinatura válida sem a `private_key.pem`.

---

## 14. Proteção contra retrocesso de data

O sistema registra a última execução válida e verifica se a data/hora atual aparenta ter sido retrocedida de forma incompatível com a execução anterior.

Finalidade:

- reduzir fraudes simples em licenças temporárias por alteração manual do relógio do Windows.

Essa proteção é complementar. A segurança principal do licenciamento continua sendo a assinatura criptográfica.

---

## 15. Fluxo operacional recomendado para o desenvolvedor

### 15.1. Recebimento do pedido

O cliente envia:

- `activation_request.json`, preferencialmente;
- ou, em último caso, a `machine_id` informada manualmente.

### 15.2. Decisão de licenciamento

O desenvolvedor define:

- nome do licenciado;
- escopo (`person`, `sector`, `company`);
- tipo de licença (`temporary` ou `perpetual`);
- data de expiração, quando aplicável;
- vínculo ou não por máquina.

### 15.3. Emissão

O desenvolvedor executa `emitir_licenca.py` com os parâmetros adequados.

### 15.4. Arquivamento

Recomenda-se guardar:

- o `activation_request.json` recebido;
- a licença gerada;
- o `license_id` utilizado;
- a data da emissão;
- eventuais observações comerciais ou técnicas.

### 15.5. Entrega

A licença gerada é enviada ao cliente.

### 15.6. Instalação

O cliente instala a licença pelo botão **Instalar licença** no próprio programa.

---

## 16. Padrão operacional recomendado

### 16.1. Licenças temporárias

Uso sugerido:

- testes;
- regularização;
- uso provisório;
- renovações curtas.

Configuração sugerida:

- sem vínculo por máquina na maioria dos casos;
- validade curta: 7, 15, 30 ou 90 dias.

### 16.2. Licenças vitalícias

Uso sugerido:

- clientes definitivos;
- liberações formais e permanentes.

Configuração sugerida:

- preferencialmente com vínculo por máquina.

### 16.3. Licenças por setor compartilhado

Há duas abordagens possíveis:

- emitir uma licença por máquina do setor;
- ou emitir licença sem `machine_id`, caso a estratégia comercial permita maior flexibilidade.

---

## 17. Itens que nunca devem ser enviados ao cliente

Nunca devem acompanhar a distribuição:

- `private_key.pem`
- scripts privados de emissão
- ferramentas internas de geração de chaves
- histórico interno de licenças emitidas

O cliente só deve receber o necessário para executar o sistema e instalar a licença.

---

## 18. Informações do build relacionadas à licença

O build do sistema deve garantir que a **chave pública** esteja acessível ao executável. O projeto já foi ajustado para procurar `public_key.pem` em mais de um local, inclusive dentro de `_internal`, quando necessário.

Mesmo assim, toda revisão de build deve confirmar:

- presença do executável;
- disponibilidade da chave pública;
- funcionamento do fluxo de instalação da licença no executável final.

---

## 19. Checklist operacional resumido

### 19.1. Antes de emitir

- [ ] confirmar se as chaves existem;
- [ ] verificar se o pedido de ativação recebido corresponde ao cliente correto;
- [ ] decidir o tipo de licença;
- [ ] decidir o escopo da licença;
- [ ] decidir se haverá vínculo por máquina.

### 19.2. Ao emitir

- [ ] gerar `license_id` único;
- [ ] preencher nome e escopo corretamente;
- [ ] definir validade se a licença for temporária;
- [ ] usar `--bind-to-activation-machine` quando necessário;
- [ ] salvar a licença gerada com nome identificável.

### 19.3. Após emitir

- [ ] arquivar a licença gerada;
- [ ] arquivar o pedido de ativação correspondente;
- [ ] registrar a entrega;
- [ ] orientar o cliente a instalar pelo botão **Instalar licença**.

---

## 20. Próximos documentos recomendados

Após este README, os documentos internos mais úteis são:

1. checklist formal de testes do licenciamento;
2. documentação do ambiente de build;
3. termo de licença de uso do software;
4. política interna de renovação e arquivamento de licenças.

---

## 21. Responsável técnico

**Desenvolvedor responsável:** Gilleade Gonçalves Silva  
**E-mail de contato:** gilleadesilva@gmail.com

