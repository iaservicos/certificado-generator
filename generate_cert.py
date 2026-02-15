# -----------------------------------------------------------------------------
# SCRIPT: generate_cert.py
# DESCRIÇÃO: Recebe dados JSON, preenche um template .docx e salva o resultado.
# -----------------------------------------------------------------------------

import sys
import json
from docxtpl import DocxTemplate

def gerar_certificado():
    """
    Função principal para gerar o certificado.
    """
    try:
        # 1. Carregar o template do Word.
        # O script espera que 'template.docx' esteja na mesma pasta que ele.
        doc = DocxTemplate("template.docx")

        # 2. Ler os dados enviados pelo Power Automate.
        # Os dados vêm como uma string JSON no primeiro argumento da linha de comando.
        # Ex: python generate_cert.py '{"nome_completo": "João da Silva", "data_conclusao": "15/02/2026", "id_resposta": 123}'
        if len(sys.argv) < 2:
            raise ValueError("Erro: Nenhum dado JSON foi fornecido.")
        
        context_data_json = sys.argv[1]
        contexto = json.loads(context_data_json)

        # 3. Renderizar o template.
        # A biblioteca substitui as variáveis {{...}} pelos valores do dicionário 'contexto'.
        doc.render(contexto)

        # 4. Definir um nome de arquivo único para o certificado gerado.
        # Usamos o ID da resposta do Forms para garantir que cada arquivo seja único.
        id_resposta = contexto.get("id_resposta", "sem_id")
        nome_arquivo_saida = f"certificado_preenchido_{id_resposta}.docx"

        # 5. Salvar o novo documento.
        doc.save(nome_arquivo_saida)

        # 6. Imprimir o nome do arquivo gerado.
        # O Power Automate vai capturar esta saída para saber qual arquivo pegar.
        print(nome_arquivo_saida)

    except Exception as e:
        # Em caso de erro, imprime o erro para o Power Automate poder registrar.
        print(f"Erro ao gerar certificado: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    gerar_certificado()

