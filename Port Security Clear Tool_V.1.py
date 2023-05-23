import concurrent.futures
import PySimpleGUI as sg
from netmiko import ConnectHandler
import requests 
import json
from charset_normalizer import md__mypyc

sg.theme("Kayak")

device_data = []
layout = [[sg.Text("Enter device details: ")],
          [sg.Table(values=device_data, headings=["Hostname/IP", "Ports to clear"], max_col_width=40, key="-TABLE-")],
          [sg.Text("Hostname/IP: "), sg.InputText(key="-IP-")],
          [sg.Text('Task Number:'), sg.InputText(key='task_num')],
          [sg.Text("Ports: Seperate the ports by commas if you need to refresh port security on multiple ports for a device. "), sg.InputText(key="-PORTS-")],
          [sg.Text("Username: "), sg.InputText(key="-USER-")],
          [sg.Text("Password: "), sg.InputText(key="-PASSWORD-", password_char="*")],
          [sg.Button("Add Device"), sg.Button("Done"), sg.Button("Clear")]]
          
window = sg.Window("Port Security Clear Tool v1", layout)

def bounce_ports(device_ip, device_port, username, password):
    output = ""
    device = {
        "device_type": "cisco_ios",
        "ip": device_ip,
        "username": username,
        "password": password,
    }
    try:
        connection = ConnectHandler(**device)
        connection.enable()
        output += (f"Connected to device: {device_ip}\n")
    except Exception as e:
        output += (f"Failed to connect to device {device_ip}: {e}\n")
        return output

    ports = device_port.split(',')
    for port in ports:
        try:
            output += (f"Shutting down port {port} on device {device_ip}\n")
            connection.send_command("config t", expect_string="#")
            connection.send_command(f"int {port}", expect_string="#")
            connection.send_command("sh", expect_string="#")
            output += (f"Clearing port security on {port} for device {device_ip}\n")
            connection.send_command("no switchport port-security mac-address sticky", expect_string="#")
            connection.send_command("no switchport port-security", expect_string="#")
            connection.send_command("switchport port-security mac-address sticky", expect_string="#")
            connection.send_command("switchport port-security", expect_string="#")
            connection.send_command("exit", expect_string="#")
            output += (f"Bringing up port {port} for device {device_ip}\n")
            connection.send_command(f"int {port}", expect_string="#")
            connection.send_command("no sh", expect_string="#")

        except Exception as e:
            output += (f"Failed to clear port security on {port} for device {device_ip}: {e}\n")
    
    output += ("The script has ran successfully!\n")
    connection.disconnect()
    return output

def run_bouncing():
    output = ""
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = [executor.submit(bounce_ports, device[0], device[1], device[2], device[3]) for device in device_data]
        for future in concurrent.futures.as_completed(results):
            output += future.result()
    return output

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, None):
        break
    if event == "Add Device":
        if not values["-IP-"]:
            sg.popup("Hostname/IP cannot be blank!")
            continue
        if not values["-PORTS-"]:
            sg.popup("Ports cannot be blank!")
            continue
        if not values["-USER-"]:
            sg.popup("Username cannot be blank!")
            continue
        if not values["-PASSWORD-"]:
            sg.popup("Password cannot be blank!")
            continue
        device_data.append([values["-IP-"], values["-PORTS-"], values["-USER-"], values["-PASSWORD-"], values["task_num"]])
        window["-TABLE-"].update(values=device_data)
    if event == "Clear":
        device_data = []
        window["-TABLE-"].update(values=[])
    if event == "Done":
        result = run_bouncing()
        result_window = sg.Window("Result", [[sg.Multiline(result, size=(80, 20), key="-RESULT-")]])
        result_window.read()
        result_window.close()
        teams_webhook_url = "https://cisgov.webhook.office.com/webhookb2/6ea28ae6-a0ad-461b-818b-a3210e45057f@5e41ee74-0d2d-4a72-8975-998ce83205eb/IncomingWebhook/0c6c92e69a1e412089a5d15fc713cef0/a9dbf621-0e82-415b-bdea-5f692fcf983a"
        headers = {
        "Content-Type": "application/json"
        }
        task_num = device_data[0][4]
        devices_message = ""
        for device in device_data:
            devices_message += "Device: " + device[0] + "," "\nPort: " + device[1] + "\n"
        data = {
        "@type": "MessageCard",
        "@context": "https://schema.org/extensions",
        "summary": "Script Completed",
        "themeColor": "0078D7",
        "title": "Port security clear action completed for " + task_num,
        "text": "Please verify my work. I ran my job successfully on the following devices and ports:\n" + devices_message
        }
        response = requests.post(teams_webhook_url, headers=headers, json=data)
        if response.status_code != 200:
         sg.popup("Failed to send message to Teams channel")
         continue
        else:
         sg.popup("Message successfully sent to Teams channel")
         continue
window.close
