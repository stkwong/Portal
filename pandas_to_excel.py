import pandas as pd

def save_cluster_xlsx(filepath, de_results, cluster_names):
    writer = pd.ExcelWriter(filepath, engine="xlsxwriter")
    for i, x in enumerate(cluster_names):
        print (i)
        print (de_results[i])
        de_results[i].to_excel(writer, sheet_name=str(x))
    writer.close()
