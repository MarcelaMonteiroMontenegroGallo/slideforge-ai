# SlideForge AI: como construí um gerador de apresentações com Amazon Bedrock e ECS Fargate

**Por Marcela Monteiro Montenegro Gallo**
*Arquiteta de Dados e AI | 9x AWS Certified | 2x Databricks Certified*
*Ingram Micro Cloud — AWS Partner*

---

## Introdução

Todo empresa tem esse momento: você tem o conteúdo da proposta no Word e tem um modelo de slides padrão, e precisa apresentação para amanhã, e você vai passar as próximas 3 horas formatando slides, ajustando cores, colocando logo no lugar certo.

Multiplica isso por 10 vezes por mês. São 30 horas de trabalho que não agregam valor nenhum ao conteúdo.

O SlideForge AI resolve isso. Você cola o texto do Word, define as cores da empresa, e em 30 segundos tem um PowerPoint formatado, com identidade visual, estruturado pela IA. Disponível em uma URL pública, acessível de qualquer lugar.

Neste artigo vou mostrar como construí essa solução do zero de exemplo para coisas do dia dia usando Amazon Bedrock, ECS Fargate, S3 e Flask, com toda a infraestrutura como código via CloudFormation.

---

## O que o SlideForge AI faz

A aplicação tem três funcionalidades principais:

**1. Upload de templates da empresa**
Você faz upload do seu template PPTX, DOCX ou PDF com as diretrizes visuais da empresa. O sistema armazena no S3 e usa como referência para manter a identidade visual.

**2. Geração de conteúdo com Amazon Bedrock**
Você cola o briefing, e-mail ou documento Word no campo de texto. O Claude 3.5 Sonnet analisa o conteúdo e estrutura automaticamente em slides com tipos específicos: capa, seções, conteúdo, métricas e encerramento.

**3. Geração do PowerPoint formatado**
A biblioteca python-pptx constrói o arquivo com as cores da empresa, tipografia consistente, numeração de slides, rodapé e layout profissional. O arquivo é salvo no S3 e uma URL pré-assinada é gerada para download imediato.

---

## Arquitetura

```
Usuário (browser)
      ↓
Application Load Balancer (URL pública)
      ↓
ECS Fargate (Flask + python-pptx)
      ↓
Amazon Bedrock (Claude 3.5 Sonnet)
      ↓
Amazon S3 (templates + outputs)
```

A escolha do ECS Fargate foi deliberada: sem gerenciar servidores, escala automática, e o custo é zero quando não há requisições (com Fargate Spot). Para uma ferramenta interna de consultoria, isso é ideal.

---

## Como o Bedrock estrutura o conteúdo

O prompt enviado ao Claude é específico sobre o formato de saída esperado:

```python
prompt = f"""Crie uma estrutura de {num_slides} slides para: {instruction}

Empresa: {company}
Cliente: {client}
Briefing: {briefing}

Retorne APENAS um JSON com esta estrutura:
{{
  "slides": [
    {{
      "type": "cover|section|content|bullets|metrics|closing",
      "title": "Título",
      "content": "Texto principal",
      "bullets": ["ponto 1", "ponto 2"],
      "metrics": [{{"label": "KPI", "value": "42", "unit": "%"}}]
    }}
  ]
}}"""
```

O Claude retorna um JSON estruturado com tipos de slide específicos. Cada tipo tem um layout diferente no python-pptx: a capa tem fundo escuro com barra lateral colorida, os slides de métricas têm cards com números em destaque, o encerramento tem CTA com próximos passos.

---

## Construção do PowerPoint com python-pptx

A parte mais interessante é a geração do PPTX. Em vez de usar templates PPTX como base (que limitam a flexibilidade), construo cada slide do zero com formas, caixas de texto e cores programáticas:

```python
def _build_cover_slide(slide, data, company, client, primary, accent, white):
    W, H = Inches(13.33), Inches(7.5)

    # Fundo com cor primária da empresa
    _add_shape_rect(slide, 0, 0, W, H, primary)

    # Barra lateral com cor de destaque
    _add_shape_rect(slide, 0, 0, Inches(0.5), H, accent)

    # Linha decorativa
    _add_shape_rect(slide, Inches(0.8), Inches(3.2), Inches(8), Inches(0.04), accent)

    # Título principal
    _add_text(slide, data.get("title", ""),
              Inches(0.9), Inches(1.5), Inches(10), Inches(1.5),
              40, white, bold=True)
```

Isso garante que qualquer combinação de cores funcione bem, porque o layout é construído programaticamente em vez de depender de um template fixo.

---

## Deploy na AWS: passo a passo

### 1. Build e push da imagem Docker

```bash
# Build
docker build -t slideforge-ai .

# Tag e push para ECR
aws ecr create-repository --repository-name slideforge-ai
docker tag slideforge-ai:latest ACCOUNT.dkr.ecr.us-east-1.amazonaws.com/slideforge-ai:latest
docker push ACCOUNT.dkr.ecr.us-east-1.amazonaws.com/slideforge-ai:latest
```

### 2. Deploy do CloudFormation

```bash
aws cloudformation deploy \
  --template-file cfn/infrastructure.yaml \
  --stack-name slideforge-prod \
  --parameter-overrides \
    ProjectName=slideforge-ai \
    Environment=prod \
    VpcId=vpc-XXXXXXXX \
    SubnetIds=subnet-XXXX,subnet-YYYY \
    ContainerImage=ACCOUNT.dkr.ecr.us-east-1.amazonaws.com/slideforge-ai:latest \
  --capabilities CAPABILITY_NAMED_IAM
```

### 3. Acessar a URL

Após o deploy, o CloudFormation retorna a URL do ALB:

```
http://slideforge-prod-alb-XXXXXXXXX.us-east-1.elb.amazonaws.com
```

Essa é a URL pública da aplicação. Qualquer pessoa com o link pode acessar e gerar apresentações.

---

## Empresa fictícia de demonstração: DataVision Consultoria

Para demonstrar o sistema, criei uma empresa fictícia chamada **DataVision Consultoria** com as seguintes configurações:

- Cor primária: `#1a3a5c` (azul marinho)
- Cor de destaque: `#ff6b00` (laranja)
- Tipo de documento: Proposta Comercial
- Cliente: TechRetail S.A.

O arquivo `sample/proposta-datavision.txt` contém uma proposta completa que você pode colar no sistema para ver o resultado. Em 30 segundos, o SlideForge gera um PowerPoint de 10 slides com:

- Slide de capa com título, cliente e empresa
- Slide de contexto e desafio
- Slide de solução proposta
- Slide de métricas com ROI e prazos
- Slides de benefícios e investimento
- Slide de próximos passos com CTA

---

## Custo estimado na AWS

Para uso interno de uma consultoria com ~50 apresentações por mês:

| Serviço | Custo estimado |
|---|---|
| ECS Fargate (1 task, ~2h/dia) | ~$8/mês |
| Amazon Bedrock (50 gerações) | ~$3/mês |
| S3 (templates + outputs) | ~$1/mês |
| ALB | ~$16/mês |
| **Total** | **~$28/mês** |

Para reduzir custo, use Fargate Spot (70% mais barato) e configure o serviço ECS para escalar para zero quando não há uso.

---

## Conclusão

O SlideForge AI não é um produto. É uma demonstração de como serviços AWS gerenciados podem resolver problemas reais de produtividade com pouco código e custo mínimo.

Amazon Bedrock para inteligência, ECS Fargate para execução sem servidor, S3 para armazenamento, ALB para exposição pública. Quatro serviços, uma URL, um problema resolvido.

O código completo está disponível no GitHub. Em menos de 30 minutos você tem o ambiente rodando na sua conta AWS.

---

*Marcela Monteiro Montenegro Gallo é Arquiteta de Dados e AI na Ingram Micro Cloud, 9x AWS Certified e 2x Databricks Certified.*
*LinkedIn: [linkedin.com/in/marcelagallo](https://linkedin.com/in/marcelagallo)*
