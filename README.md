# PPTX

## Documentação

https://python-pptx.readthedocs.io/en/latest/index.html

## Instalação

`pip install python-pptx`

## Importação

`from pptx import Presentation`

## Criando/Editando/Salvando apresentação

```py
from pptx import Presentation

# criando
apresentacao = Presentation()

# editando


#salvando
apresentacao.save("MeuPPT.pptx")
```
## Criando slides baseado em layouts padrões

```py
from pptx import Presentation

# criando
apresentacao = Presentation()

# editando (adicionar um elemento => editar elemento)
slide1 = apresentacao.slides.add_slide(apresentacao.slide_layouts[0])

#salvando
apresentacao.save("MeuPPT.pptx")
```

### Inserindo elementos em slides com layouts padrões

**Observação:** SEMPRE quando quisermos inserir um elemento no slide, nós primeiro adicionar o elemento e depois editamos ele

```py
from pptx import Presentation

# criando
apresentacao = Presentation()

# editando (adicionar um elemento => editar elemento)
slide1 = apresentacao.slides.add_slide(apresentacao.slide_layouts[0])

"""
Existe dois formas de adicionar elementos no slide:

* Acessando os placeholders
* Acessando os shapes

Placeholders ou shapes são as caixas que já vem por padrão no seu slide.

No exemplo abaixo acessamos o shape chamada "subtitle", porém para acessarmos desta forma o shape DEVE ser um shape nomeado por padrão ou ser nomeado manualmente.
"""

titulo = slide1.placeholders[0] 
subtitulo = slide1.shapes.subtitle

titulo.text = "1º Slide"
subtitulo.text = "Estamos criando um ppt com Python"

#salvando
apresentacao.save("MeuPPT.pptx")
```

## Criando slides personalizados

```py
from pptx import Presentation
from pptx.util import Inches, Pt # Polegadas

# criando
apresentacao = Presentation()

# editando
slide = apresentacao.slides.add_slide(apresentacao.slide_layouts[6]) # slide em branco

x = Inches(1)
y = Inches(1)
lar = Inches(2)
alt = Inches(2)

caixa_texto = slide.shapes.add_textbox(x, y, larg, alt)

# 1 forma de editar uma caixa de texto
caixa_texto.text = "Vendas de Janeiro"

# 2 forma de editar uma caixa de texto
text_fram = caixa_texto.text_frame
paragrafo = text_fram.add_paragraph()
paragrafo.text = "R$ 10.000"
paragrafo.font.bold = True
paragrafo.font.size = Pt(30)
paragrafo.font.color = ""

#salvando
apresentacao.save("MeuPPT.pptx")
```

## Criando slides com gráficos

```py
from pptx import Presentation
from pptx.util import Inches, Pt # Polegadas
from pptx.chart.data import CategoryChartData # Possiblidade de add categoria no gráfico
from pptx.enum.chart import XL_CHART_TYPE # Tipo de gráfico

# criando
apresentacao = Presentation()

# editando
slide = apresentacao.slides.add_slide(apresentacao.slide_layouts[6]) # slide em branco

# criar gráfico
produtos = ["Iphone", "IPad", "Airpod"] # categorias (eixo x)
vendas = [1500, 1000, 2000] # serie de dados (eixo y)

x = Inches(1)
y = Inches(1)
lar = Inches(5)
alt = Inches(3)

dados_grafico = CategoryChartData()
dados_grafico.categories = produtos
dados_grafico.add_series("Vendas", vendas)

slide.shapes.add_chart(XL_CHART_TYPE.NOME_DO_TIPO, x, y, larg, alt, dados_grafico)

#salvando
apresentacao.save("MeuPPT.pptx")
```
