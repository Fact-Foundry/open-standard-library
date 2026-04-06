using OslSpreadsheet.Models;
using OslSpreadsheet.Models.Files.ods;
using System.Text;

namespace OslSpreadsheet.Services
{
    internal class OdsFileService : IFileService
    {
        private bool disposedValue;

        public async ValueTask DisposeAsync()
        {
            await Task.Run(() => Dispose());
        }

        public async Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            byte[] output;

            try
            {
                var content = GenerateContentFile(workbook);
                var meta = GenerateMetaFile(workbook);
                var style = GenerateStyleFile(workbook);

                // Add models to a file list for compression
                List<InMemoryFile> files = new()
                {
                    new InMemoryFile()
                    {
                        FileName = "META-INF/manifest.xml",
                        Content = await XmlService.ConvertToXmlAsync(new ODManifest())
                    },
                    new InMemoryFile()
                    {
                        FileName = "content.xml",
                        Content = await XmlService.ConvertToXmlAsync(content)
                    },
                    new InMemoryFile()
                    {
                        FileName = "meta.xml",
                        Content = await XmlService.ConvertToXmlAsync(meta)
                    },
                    new InMemoryFile()
                    {
                        FileName = "mimetype",
                        Content = Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.spreadsheet")
                    },
                    new InMemoryFile()
                    {
                        FileName = "styles.xml",
                        Content = await XmlService.ConvertToXmlAsync(style)
                    }
                };

                output = await ZipService.GenerateZipAsync(files);
            }
            catch
            {
                throw new Exception("There was an issue compressing the file.");
            }

            return output;
        }

        public async Task<oWorkbook> GenerateModel(byte[] file)
        {
            throw new NotImplementedException();
        }

        private ODContent GenerateContentFile(oWorkbook workbook)
        {
            var file = new ODContent();

            foreach (var s in workbook.Sheets)
            {
                var masterPageName = string.Format("mp{0}", s.Index);
                var styleName = string.Format("ta{0}", s.Index);

                file.automaticStyles.automaticStyles.Add(new ODContent.AutomaticStyles.Style()
                {
                    Name = styleName,
                    Family = "table",
                    MasterPageName = masterPageName,
                    tableProperties = new()
                });

                var table = new ODContent.Table()
                {
                    Name = s.SheetName,
                    StyleName = styleName
                };

                if (s.Cells.Any())
                {
                    int rowCount = s.RowCount;
                    int colCount = s.ColumnCount;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        var tableRow = new ODContent.Table.TableRow();

                        for (int c = 1; c <= colCount; c++)
                        {
                            var cell = s.Cells.FirstOrDefault(x => x.Row == r && x.Column == c);

                            if (cell != null)
                            {
                                var tableCell = new ODContent.Table.TableRow.TableCell()
                                {
                                    StyleName = "ce1",
                                    TextValue = cell.Value
                                };

                                if (cell.ValueType == CellValueType.Float)
                                {
                                    tableCell.ValueType = "float";
                                    tableCell.NumericValue = cell.Value;
                                }
                                else
                                {
                                    tableCell.ValueType = "string";
                                }

                                tableRow.Cells.Add(tableCell);
                            }
                            else
                            {
                                tableRow.Cells.Add(new ODContent.Table.TableRow.TableCell());
                            }
                        }

                        // Trailing empty columns filler
                        tableRow.Cells.Add(new ODContent.Table.TableRow.TableCell()
                        {
                            NumberColumnsRepeated = (16384 - colCount).ToString()
                        });

                        table.Rows.Add(tableRow);
                    }

                    // Trailing empty rows filler
                    table.Rows.Add(new ODContent.Table.TableRow()
                    {
                        NumberRowsRepeated = (1048577 - rowCount).ToString(),
                        Cells = new List<ODContent.Table.TableRow.TableCell>()
                        {
                            new ODContent.Table.TableRow.TableCell()
                            {
                                NumberColumnsRepeated = "16384"
                            }
                        }
                    });
                }
                else
                {
                    // Empty sheet — filler row
                    table.Rows.Add(new ODContent.Table.TableRow()
                    {
                        NumberRowsRepeated = "1048577",
                        Cells = new List<ODContent.Table.TableRow.TableCell>()
                        {
                            new ODContent.Table.TableRow.TableCell()
                            {
                                NumberColumnsRepeated = "16384"
                            }
                        }
                    });
                }

                file.body.spreadsheet.Tables.Add(table);
            }

            return file;
        }

        /// <summary>
        /// Generates meta.xml file found in the root of the ODS zip file
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private ODMeta GenerateMetaFile(oWorkbook workbook)
        {
            return new ODMeta()
            {
                meta = new ODMeta.Meta()
                {
                    CreationDate = workbook.CreationDate,
                    Creator = workbook.Creator,
                    Date = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ"),
                    Generator = workbook.Generator
                }
            };
        }

        private ODStyles GenerateStyleFile(oWorkbook workbook)
        {
            var file = new ODStyles();

            foreach (var s in workbook.Sheets)
            {
                var masterPageName = string.Format("mp{0}", s.Index);
                var pageLayoutName = string.Format("pm{0}", s.Index);

                file.automaticStyles.pageLayout.Add(new ODStyles.AutomaticStyles.PageLayout()
                {
                    Name = pageLayoutName
                });

                file.masterStyles.masterPage.Add(new ODStyles.MasterStyles.MasterPage()
                {
                    FooterLeftPageStyle = new ODStyles.MasterStyles.MasterPage.PageStyle()
                    {
                        Display = "false"
                    },
                    HeaderLeftPageStyle = new ODStyles.MasterStyles.MasterPage.PageStyle()
                    {
                        Display = "false"
                    },
                    Name = masterPageName,
                    PageLayoutName = pageLayoutName
                });
            }

            return file;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~GenerateOdsFileService()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
