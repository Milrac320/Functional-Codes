# Readme - Validação de CPF no Excel com VBA

Este projeto tem como objetivo criar uma função personalizada no Excel, utilizando VBA (Visual Basic for Applications), para validar um CPF (Cadastro de Pessoa Física) de forma automática.
A função verifica se o CPF inserido é válido de acordo com as regras de cálculo dos dígitos verificadores estabelecidas pela Receita Federal do Brasil.

## Como funciona a validação de CPF

O CPF é composto por 11 dígitos, onde os dois últimos dígitos são chamados de dígitos verificadores. Estes dígitos são calculados com base nos 9 primeiros dígitos do CPF, usando um algoritmo específico que segue as seguintes etapas:

***Primeiro dígito verificador:***

- Multiplica-se cada um dos primeiros 9 dígitos pela sequência de números decrescentes de 10 a 2.

- Soma-se o resultado dessas multiplicações.

- O resultado dessa soma é dividido por 11, e o resto da divisão é subtraído de 11.

- Se o resultado for maior que 9, o primeiro dígito verificador é 0; caso contrário, é o próprio resultado.

***Segundo dígito verificador:***

- Multiplica-se cada um dos primeiros 10 dígitos (incluindo o primeiro dígito verificador) pela sequência de números decrescentes de 11 a 2.

- Soma-se o resultado dessas multiplicações.

- O resultado dessa soma é dividido por 11, e o resto da divisão é subtraído de 11.

- Se o resultado for maior que 9, o segundo dígito verificador é 0; caso contrário, é o próprio resultado.

***Verificação final:***

Compara-se os dois dígitos verificadores calculados com os dois dígitos verificadores originais do CPF.
Se ambos forem iguais, o CPF é considerado válido. Caso contrário, é inválido.

## Requisitos

- Microsoft Excel (com suporte a macros habilitado)

- Editor VBA para Excel (embutido no Excel)

## Considerações importantes

- Formato do CPF: A função remove automaticamente pontos (.) e traços (-), então o CPF pode ser inserido com ou sem esses caracteres.

- Tamanho do CPF: A função verifica se o CPF possui 11 dígitos, o formato correto.

- Dígitos repetidos: CPFs formados por números repetidos (como 111.111.111-11) são considerados inválidos, conforme as regras da Receita Federal.

## Fundamentos do Algoritmo

O algoritmo de validação de CPF utiliza multiplicações, somas e operações modulares baseadas nos nove primeiros dígitos do CPF. Este método é utilizado pela Receita Federal para assegurar que um CPF seja autêntico.
Os dígitos verificadores (os dois últimos) são calculados a partir dos dígitos anteriores, e apenas CPFs que seguem esse padrão são considerados válidos.

***Exemplo***

Considere o CPF 123.456.789-09. Os passos para validar este CPF são:
Multiplicando os primeiros 9 dígitos por 10 a 2 e somando:

***(1∗10)+(2∗9)+(3∗8)+(4∗7)+(5∗6)+(6∗5)+(7∗4)+(8∗3)+(9∗2)=210***

O resto da divisão de 210 por 11 é 1, e 11 - 1 = 10. Como o resultado é maior que 9, o primeiro dígito verificador é 0.

Multiplicando os primeiros 10 dígitos por 11 a 2 e somando:

***(1∗11)+(2∗10)+(3∗9)+(4∗8)+(5∗7)+(6∗6)+(7∗5)+(8∗4)+(9∗3)+(0∗2)=237***

O resto da divisão de 237 por 11 é 6, e 11 - 6 = 5. O segundo dígito verificador é 5.

Comparando os dígitos verificadores calculados (0 e 5) com os dígitos originais (0 e 9), o CPF 123.456.789-09 é inválido.

## Autor

- Carlos Garcia
- **[@carlosgarcia.programacao](https://www.instagram.com/carlosgarcia.programacao/)**
