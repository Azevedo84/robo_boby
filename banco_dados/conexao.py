import fdb

conecta = fdb.connect(database=r'C:\HallSys\db\Horus\Suzuki\ESTOQUE.GDB',
                              host='PUBLICO',
                              port=3050,
                              user='sysdba',
                              password='masterkey',
                              charset='ANSI')


conecta_robo = fdb.connect(database=r'C:\HallSys\db\Horus\Suzuki\ROBOZINHO.GDB',
                              host='PUBLICO',
                              port=3050,
                              user='sysdba',
                              password='masterkey',
                              charset='ANSI')
