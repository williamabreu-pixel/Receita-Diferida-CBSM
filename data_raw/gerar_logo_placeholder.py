# gerar_logo_placeholder.py
"""Gera um placeholder do ícone Dotz (círculo laranja + "DZ") em PNG com
fundo transparente, para usar no header do app enquanto o arquivo oficial
da marca não é exportado. Rodar com: python gerar_logo_placeholder.py
Troque icone_dotz.png pelo arquivo oficial quando disponível — este é só
uma aproximação, não é pixel-perfect ao logo real (a curva do "Z" é
estilizada no original)."""

from PIL import Image, ImageDraw, ImageFont

TAMANHO = 512
COR_LARANJA = (255, 79, 13, 255)  # #FF4F0D
COR_PRETO = (0, 0, 0, 255)

FONTES_CANDIDATAS = [
    r"C:\Windows\Fonts\seguibl.ttf",   # Segoe UI Black
    r"C:\Windows\Fonts\segoeuib.ttf",  # Segoe UI Bold
    r"C:\Windows\Fonts\arialbd.ttf",   # Arial Bold
]


def _carregar_fonte(tamanho):
    for caminho in FONTES_CANDIDATAS:
        try:
            return ImageFont.truetype(caminho, tamanho)
        except OSError:
            continue
    return ImageFont.load_default()


def gerar(caminho_saida="icone_dotz.png"):
    img = Image.new("RGBA", (TAMANHO, TAMANHO), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([0, 0, TAMANHO, TAMANHO], fill=COR_LARANJA)

    texto = "DZ"
    fonte = _carregar_fonte(int(TAMANHO * 0.4))
    bbox = draw.textbbox((0, 0), texto, font=fonte)
    largura, altura = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pos = ((TAMANHO - largura) / 2 - bbox[0], (TAMANHO - altura) / 2 - bbox[1])
    draw.text(pos, texto, font=fonte, fill=COR_PRETO)

    img.save(caminho_saida)
    print(f"Ícone placeholder gerado: {caminho_saida}")


if __name__ == "__main__":
    gerar("icone_dotz.png")
    gerar("logo_dotz.png")
