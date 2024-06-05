dias=float(input('Quantos dias você alugou o carro? '))
km=float(input('Quantos KM rodados ? '))
tkm=km*0.15
td=dias*60
print(f'O valor do aluguel ficou em: R$ {tkm+td}')