from time import sleep

import psycopg2
from openpyxl import load_workbook





def kek():

    conn = psycopg2.connect(dbname=db_name, host='172.16.10.22', port='5432',
                            user=db_username, password=db_password)

    cur = conn.cursor(name='1583_first_part')

    query = f"""
           select qfo.source_store_id,
                  ds.sale_source_obj_id,
                  ds.store_name,
                  dss.name_1c,
                  sum(qfo.cnt_wares_pos_detail)                                                             as "Кол-во товаров в чеках", -- "Количество товаров в чеке"
                  sum(qfo.order_sum_with_vat)                                                               as "Выручка, тг с НДС",
                  sum(qfo.order_sum_without_vat)                                                            as "Выручка, тг без НДС",
                  sum(qfo.is_sale_order)                                                                    as "Количество чеков"
           from dwh_data.qs_fact_order qfo
           left join dwh_data.dim_store ds on ds.source_store_id = qfo.source_store_id and current_date between ds.datestart and ds.dateend
           left join dwh_data.dim_store_src dss on dss.source_store_id = qfo.source_store_id
           where date(qfo.order_date) between to_date('{start_date}', 'YYYY-MM-DD') and to_date('{end_date}', 'YYYY-MM-DD')
                   and qfo."source_store_id"::int > 0 and ds."store_name" like '%Торговый%'
           group by ds.sale_source_obj_id, qfo.source_store_id, ds.store_name, dss.name_1c
                union all
           select   qfog.source_store_id,
                    ds.sale_source_obj_id,
                    ds.store_name,
                    dss.name_1c,
                    sum(qfog.count_wares)                            as quantity,
                    sum(qfog.sum_sale_vat)                           as order_sum_with_vat,
                    sum(qfog.sum_sale_no_vat)                        as order_sum_without_vat,
                    sum(qfog.is_sale_order)                          as sale_order_cnt
           from dwh_data.qs_fact_order_gross qfog
           left join dwh_data.dim_store ds on ds.source_store_id = qfog.source_store_id and current_date between ds.datestart and ds.dateend
           left join dwh_data.dim_store_src dss on dss.source_store_id = qfog.source_store_id
           where date(qfog.order_date) between to_date('{start_date}', 'YYYY-MM-DD') and to_date('{end_date}', 'YYYY-MM-DD')
                 and qfog."source_store_id"::int > 0 and ds."store_name" like '%Торговый%'
           group by ds.sale_source_obj_id, qfog.source_store_id, ds.store_name, dss.name_1c
           order by source_store_id;
       """

    cur.execute(query)

    logger.info('Executed first request')

    df = pd.DataFrame(cur.fetchall())

    cur.close()
    conn.close()

# print(''.join('0' for i in range(13 - len(str(a)))), a, sep='')





