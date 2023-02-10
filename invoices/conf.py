#your_target_folder = '/Users/docha/PycharmProjects/tati_invoices'
print('union (u), union riga (r), alvigroup (нажать любую клавишу):')
who_firma = input()
if who_firma == 'u':
    your_target_folder_docha = r'C:/Users/docha/OneDrive/Leka/INVOICE'
    your_target_folder = r'C:\Users\Admin\OneDrive\Gussev\UNION\Invoices_acts\Leka\INVOICE'
elif who_firma == 'r':
    your_target_folder_docha = r'C:/Users/docha/OneDrive/Leka/INVOICE_Riga'
    your_target_folder = r'C:\Users\Admin\OneDrive\Gussev\UNION\Invoices_acts\Leka\INVOICE_Riga'
else:
    your_target_folder_docha = r'C:/Users/docha/OneDrive/Leka/INVOICE_Alvigroup'
    your_target_folder = r'C:\Users\Admin\OneDrive\Gussev\UNION\Invoices_acts\Leka\INVOICE_Alvigroup'