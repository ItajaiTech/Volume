# util.local

O Volume agora usa `util.local` como dominio interno padrao na porta `6100`.

Configuracao aplicada no projeto:

- Host publico padrao: `util.local`
- Porta padrao: `6100`
- Bind interno padrao: `0.0.0.0`

URL esperada:

- `http://util.local:6100`

Para o dominio resolver no Windows, voce precisa ter uma destas opcoes:

1. DNS interno apontando `util.local` para o IP da maquina que executa o Volume.
2. Entrada no arquivo `hosts` da propria maquina.

Exemplo de entrada em `C:\Windows\System32\drivers\etc\hosts`:

```text
127.0.0.1 util.local
```

Se o acesso for a partir de outras maquinas da rede, use o IP LAN do servidor em vez de `127.0.0.1`.

Script auxiliar incluso no projeto:

- `C:\Volume\shipping_ai\configure_util_local.ps1`

Exemplos:

```powershell
powershell -ExecutionPolicy Bypass -File C:\Volume\shipping_ai\configure_util_local.ps1
powershell -ExecutionPolicy Bypass -File C:\Volume\shipping_ai\configure_util_local.ps1 -TargetIp 192.168.1.50
```

Variaveis de ambiente suportadas:

- `VOLUME_PUBLIC_HOST`: host usado para abrir a URL e montar o `SERVER_NAME`.
- `VOLUME_BIND_HOST`: endereco em que o servidor escuta. Padrao `0.0.0.0`.
- `VOLUME_PORT`: porta do servico. Padrao `6100`.
- `VOLUME_SERVER_NAME`: sobrescreve o `SERVER_NAME` do Flask, por exemplo `util.local:6100`.
- `VOLUME_TRUSTED_HOSTS`: hosts extras separados por virgula.