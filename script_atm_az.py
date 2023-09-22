from azure.identity import DefaultAzureCredential
from azure.mgmt.compute import ComputeManagementClient
from openpyxl import Workbook
import pandas as pd
import concurrent.futures

# Variáveis de configuração
resource_group_to_ignore = ["rg-1", "rg-2"]
output_excel_file = "geracao_pd/informacoes_servidores.xlsx"
pd_base_file = "geracao_pd/pd_base.xlsx"
new_info_excel_file = "geracao_pd/informacoes_servidores.xlsx"
output_file = "geracao_pd/planilha_atualizada.xlsx"


def status_servidor(credential, subscription_id, resource_group_name, servidor):
    client = ComputeManagementClient(
        credential=credential,
        subscription_id=subscription_id,
    )

    response = client.virtual_machines.instance_view(
        resource_group_name=resource_group_name,
        vm_name=servidor,
    )

    statuses = response.statuses

    if statuses:
        primeiro_status = statuses[1]
        status_servidor = primeiro_status.display_status
        return status_servidor

    return "N/A"


def process_vm(credential, subscription_id, resource_group, servidor):
    status = status_servidor(credential, subscription_id, resource_group, servidor.name)

    if servidor.tags:
        formatted_tags = ", ".join(
            [f"{key}: {value}" for key, value in servidor.tags.items()]
        )
    else:
        formatted_tags = "N/A"

    formatted_zones = ", ".join(servidor.zones) if servidor.zones else "N/A"

    vm_size = servidor.hardware_profile.vm_size
    vcpus = extract_vcpus(vm_size)

    return [
        servidor.name,
        subscription_id,
        resource_group,
        formatted_tags,
        status,
        servidor.location,
        vm_size,
        vcpus,
        formatted_zones,
    ]


def extract_vcpus(vm_size):
    parts = vm_size.split("_")
    vcpus = parts[1] if len(parts) >= 2 else "N/A"
    return vcpus


def encontrar_linhas_novas(pd_base_file, new_info_excel_file):
    pd_base = pd.read_excel(pd_base_file, usecols=["Hostname"])["Hostname"].tolist()
    nova_planilha = pd.read_excel(new_info_excel_file, usecols=["Hostname"])[
        "Hostname"
    ].tolist()

    linhas_novas = set(nova_planilha) - set(pd_base)

    if linhas_novas:
        print("Linhas novas (hostname) encontradas:")
        for hostname in linhas_novas:
            print(hostname)
    else:
        print("Nenhuma linha nova encontrada.")


def adicionar_linhas(pd_base_file, new_info_excel_file, output_file):
    pd_base = pd.read_excel(pd_base_file)
    nova_planilha = pd.read_excel(new_info_excel_file)

    hostnames_faltantes = set(nova_planilha["Hostname"]) - set(pd_base["Hostname"])
    novas_linhas = nova_planilha[nova_planilha["Hostname"].isin(hostnames_faltantes)]

    resposta_user = input("Quer gerar uma nova planilha atualizada? Y ou N: ").lower()
    if resposta_user == "y":
        pd_base = pd.concat([pd_base, novas_linhas], ignore_index=True)
        pd_base.to_excel(output_file, index=False)
        print(f"{len(novas_linhas)} novas linhas adicionadas a {output_file}")
    else:
        print("Não foi gerada nenhuma planilha.")


credential = DefaultAzureCredential()

subscriptions = [
    "sub-1",
    "sub-2",
    "sub-3",
    "sub-4",
]

workbook = Workbook()
sheet = workbook.active
sheet.append(
    [
        "Hostname",
        "Subscription",
        "Resource Group",
        "Tags",
        "Status",
        "Localizacao",
        "VM Size",
        "vCPus",
        "Zona",
    ]
)

with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = []
    for subscription_id in subscriptions:
        compute_client = ComputeManagementClient(credential, subscription_id)
        servidores = compute_client.virtual_machines.list_all()
        for servidor in servidores:
            resource_id_parts = servidor.id.split("/")
            resource_group_index = resource_id_parts.index("resourceGroups") + 1
            resource_group = resource_id_parts[resource_group_index]
            if resource_group in resource_group_to_ignore:
                continue
            subscription_index = resource_id_parts.index("subscriptions") + 1
            subscription = resource_id_parts[subscription_index]
            futures.append(
                executor.submit(
                    process_vm, credential, subscription, resource_group, servidor
                )
            )
    for future in concurrent.futures.as_completed(futures):
        sheet.append(future.result())

workbook.save(output_excel_file)

print(f"Arquivo salvo com sucesso {output_excel_file}")
print()
print("Iniciando comparativo de planilhas...")
print()

encontrar_linhas_novas(pd_base_file, new_info_excel_file)
adicionar_linhas(pd_base_file, new_info_excel_file, output_file)