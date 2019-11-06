const { google } = require('googleapis');

/**
 * Exemplos:
 * 
 * Insere dados na aba que nasce com uma planilha 
 *  const data = [['Est 1', 'teste', 2250, '2019-12-25'], ['Est 2', 'teste 2', 10350, '2019-12-14T14:10:20.000Z']]
 *  await sheet.addRows(0, data);
 * 
 * Formata o visual das colunas C e D
    await sheet.changeFormat(0, 2, ['INTEGER','BOLD']);
    await sheet.changeFormat(0, 3, ['BR_DATE_TIME', 'BOLD']);

   Insere um cabeçalho com títulos para cada coluna na aba 211909232
    const dataHeader = ['Estação', 'Texto', 'Metrica', 'Data Publicação'];
    await sheet.addHeader(211909232, dataHeader);
    
   Insere dados
    const data = [['Est 1', 'teste', 2250, '2019-12-25'], ['Est 2', 'teste 2', 10350, '2019-12-14T14:10:20.000Z']]
    await sheet.addRows(211909232, data);
    
   Formata
    await sheet.changeFormat(211909232, 2, ['BOLD','INTEGER']);
    await sheet.changeFormat(211909232, 3, ['BR_DATE']);

   Formats
    'INTEGER','BOLD','SPECIAL_DATE','BR_DATE','BR_DATE_TIME'
    
 */
class GoogleSheet {
    
    /**
     * Inicializa acesso a uma planilha
     * 
     * @param {*} credentials       Dados de autenticação
     * @param {*} spreadsheetId     ID da planilha
     */
    constructor(credentials, spreadsheetId) {
        const auth = new google.auth.JWT(
            credentials.client_email,
            null,
            credentials.private_key,
            'https://www.googleapis.com/auth/spreadsheets'
        );

        this.sheets = google.sheets({ version: 'v4', auth });
        this.spreadsheetId = spreadsheetId;
    }

    /**
     * Adiciona um header em uma aba da planilha
     * 
     * @param {*} sheetTabId ID da aba
     * @param {*} headNames  Títulos de cada coluna
     * @return {Void}
     */
    async addHeader(sheetTabId, headNames) {
        try {
            await this.sheets.spreadsheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [
                        {
                            appendCells: {
                                sheetId: sheetTabId,
                                rows: this._dataFormat([headNames]),
                                fields: '*'
                            },
                        },
                        {
                        repeatCell: {
                        range: {
                            sheetId: sheetTabId,
                            startRowIndex: 0,
                            endRowIndex: 1
                        },
                        cell: {
                            userEnteredFormat: {
                            horizontalAlignment : "CENTER",
                            textFormat: {
                                fontSize: 12,
                                bold: true
                            }
                            }
                        },
                        fields: "userEnteredFormat(textFormat,horizontalAlignment)"
                        }
                    },
                    {
                        updateSheetProperties: {
                        properties: {
                            sheetId: sheetTabId,
                            gridProperties: {
                            frozenRowCount: 1
                            }
                        },
                        fields: "gridProperties.frozenRowCount"
                        }
                    }
                    ]
                }
            });
        } catch(e) {
            throw new Error(e.message);
        }
    }

    /**
     * Atualiza a planilha
     * 
     * @param {*} sheetTabId    ID da aba
     * @param {*} data          dados a serem inseridos no final da planilha
     * @return {Void}
     */
    async addRows(sheetTabId, data) {
        try {
            await this.sheets.spreadsheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [
                        {
                            appendCells: {
                                sheetId: sheetTabId,
                                rows: this._dataFormat(data),
                                fields: '*'
                            },
                        },
                    ]
                }
            });
        } catch(e) {
            throw new Error(e.message);
        }
    }

    /**
     * Atualiza formatação
     * 
     * @param {*} sheetTabId    ID da aba
     * @param {*} column        Coluna que receberá a formatação
     * @param {*} formats       tipos de formatação
     * @return {Void}
     */
    async changeFormat(sheetTabId, column, formats) {
        try {
            await this.sheets.spreadsheets.batchUpdate({
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [
                        {
                            repeatCell: {
                                range: {
                                    sheetId: sheetTabId,
                                    startRowIndex: 1,
                                    startColumnIndex: column,
                                    endColumnIndex: column + 1,
                                },
                                cell: this._typeFormat(formats),
                                fields: "userEnteredFormat"
                            },
                        }
                    ]
                }
            });
        } catch(e) {
            throw new Error(e.message);
        }
    }

    /**
     * Monta as linhas a serem inseridas no final da planilha
     * 
     * @param {*} data          Lista com os conteúdos de uma linha da planilha
     * @return {*}              Objeto com o dado de entrada e seu respectivo tipo
     */
    _dataFormat(data) {
        return data.map(row => {
            const localRow = row.map(item => {
                const itemType = this._typeOf(item);
                const itemValue = itemType == 'date'? this._dateFormat(item) : item;
                let sheetType = 'stringValue';

                switch(itemType) {
                    case 'string':
                        sheetType = 'stringValue';
                        break;
                    case 'number':
                        sheetType = 'numberValue';
                        break;    
                    case 'boolean':
                        sheetType = 'boolValue';
                        break;
                    case 'date':
                        sheetType = 'numberValue';
                        break;
                }

                return {
                    userEnteredValue: {
                        [sheetType]: itemValue
                    }
                }
            }, []);

            return {
                values: localRow
            }
        }, []);
    }

    /**
     * Formatação das colunas
     * 
     * @param {Array} formats   Lista com os formatos que devem ser aplicados a uma célula
     * @return {*}              Objeto com as configs a serem aplicados na coluna
     */
    _typeFormat(formats = []) {
        let res = {
            userEnteredFormat: {}
        };

        formats.forEach(format => {
            switch(format) {
                case 'INTEGER':
                    Object.assign(res.userEnteredFormat, {
                        numberFormat: {
                            type: "NUMBER",
                            pattern: "#,#0",
                        }
                    });
                    break;
                case 'BOLD':
                    Object.assign(res.userEnteredFormat, {
                        textFormat: {
                            bold: true
                        }
                    });
                    break;
                case 'SPECIAL_DATE':
                    Object.assign(res.userEnteredFormat, {
                        numberFormat: {
                            type: "DATE",
                            pattern: "mmm.-yy",
                        }
                    });
                    break;
                case 'BR_DATE':
                    Object.assign(res.userEnteredFormat, {
                        numberFormat: {
                            type: "DATE",
                            pattern: "dd/mm/yyyy",
                        }
                    });
                    break;
                case 'BR_DATE_TIME':
                    Object.assign(res.userEnteredFormat, {
                        numberFormat: {
                            type: "DATE",
                            pattern: "dd/mm/yyyy hh:mm:ss",
                        }
                    });
                    break;
            }
        });

        return res;
    }

    /**
     * Converte uma data para o formato aceito pelo Google Sheet
     * 
     * @param {*} date  Data
     * @return {Number}      Decimal
     */
    _dateFormat(date = new Date().toISOString()) {
        return (Date.parse(date) - Date.parse('1899-12-30T00:00:00.000Z')) / ( 1000 * 60 * 60 * 24 )
    }

    /**
     * Define o tipo de dado inserido na célula
     * 
     * @param {*} val       Valor do dado
     * @return {String}     String
     */
    _typeOf(val) {
        if(typeof val == 'string') {
            if(val.match(/\d{4}\-\d{2}\-\d{2}/gm) !== null) {
                return 'date';
            }
        }

        return typeof val;
    }
};

module.exports = GoogleSheet;