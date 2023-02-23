import pandas.core.series
from zsd_toma import toma_prd
import pandas as pd
from datetime import datetime
from requestSQLHana import GetDataHana
import warnings
warnings.filterwarnings('ignore')


class CopiarRechazar:
    def __init__(self, ruta_excel):
        self.path_excel_pedidos = ruta_excel
        self.dia_str = datetime.today().date().strftime("%d-%m-%Y")
        self.instancia_bbdd = None
        self.df_mergeado = None

    def create_date(self):
        return datetime.now().today().strftime("%d-%m-%Y__%Hhs-%Mmin-%Ss")

    def traer_datos_excel(self):
        try:
            print(f">>> Extrayendo pedidos y solicitantes nuevos del excel")
            data_frame = pd.read_excel(self.path_excel_pedidos)
            return data_frame if not data_frame.empty else False
        except Exception as ex:
            print(f">> Se produjo una excepcion al intentar obtener datos de pedidos del excel. Ex: {ex}")
            return False

    def traer_fecha_pedidos(self, conjunto_pedidos):
        # conjunto_pedidos = tuple(map(lambda x: f"000{x}", conjunto_pedidos))
        self.instancia_bbdd = GetDataHana()
        print(f">>> Realizando consulta a BBDD para pedidos: {conjunto_pedidos}")
        if self.instancia_bbdd:
            self.instancia_bbdd.define_query(conjunto_pedidos)
            df_info_pedidos = self.instancia_bbdd.get_request_data()
            # Se extrae la columna pedido y se aplica la funcion map para quitar los ceros al inicio
            pedidos_sin_ceros_al_inicio = df_info_pedidos["PEDIDOS"].map(lambda x: int(x))
            # Reemplazar la columna pedido por la nueva serie donde se aplico la funcion map
            df_info_pedidos["PEDIDOS"] = pedidos_sin_ceros_al_inicio
            return df_info_pedidos if not df_info_pedidos.empty else False
        else:
            print(f">>> No se pudo establecer una conexion a la BBDD")

    def unir_data_frames(self, df_excel: pandas.DataFrame, df_fechas_y_pedidos: pandas.DataFrame):
        try:
            print(f">>> Intentando unir dataFrame Excel con dataFrame de pedidos y fechas de SAP")
            self.df_mergeado = df_excel.merge(df_fechas_y_pedidos, on="PEDIDOS", how="left")
            return self.df_mergeado if not self.df_mergeado.empty else False
        except Exception as ex:
            print(f">> No se pudo realizar el merge de los dataframes. Error: {ex}")

    def modificar_solicitante_rnos(self, df_pedidos_modificar):
        lista_pedidos = []
        for pedido, solicitante, fecha in zip(df_pedidos_modificar["PEDIDOS"], df_pedidos_modificar["NSOL"], df_pedidos_modificar["FE_ENTREGA"]):
            print(f"Pedido: {pedido}, NSOL: {solicitante}, Fecha: {fecha}")
            resultados = toma_prd(pedido, solicitante, fecha, 0)
            lista_pedidos.append(resultados)
            print("------------")
        df_resultados = pd.DataFrame(lista_pedidos, columns=["PEDIDOS", "MENSAJE TOMA", "MENSAJE VA02"])
        print(df_resultados)
        resultado_df_final = df_pedidos_modificar.merge(df_resultados, on="PEDIDOS", how="left")
        print(resultado_df_final)
        resultado_df_final.to_excel("../resources/pedidos_copiar_rechazar_resultados.xlsx", index=False)


objeto_cop_rech = CopiarRechazar("../resources/pedidos_copiar_rechazar.xlsx")
df_excel = objeto_cop_rech.traer_datos_excel()
pedidos_excel = df_excel["PEDIDOS"]
df_fechas_y_pedidos = objeto_cop_rech.traer_fecha_pedidos(tuple(pedidos_excel))
data_frames_unidos = objeto_cop_rech.unir_data_frames(df_excel, df_fechas_y_pedidos)
data_frames_unidos.to_excel("../resources/pedidos_copiar_rechazar_fechas.xlsx")

# Iterar pedidos
objeto_cop_rech.modificar_solicitante_rnos(data_frames_unidos)

