using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using UnityEngine;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Xml.Serialization;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using UnityEditor;

namespace E2J {
    public class E2JManager : AssetPostprocessor{

        private static readonly string tablePath = "Assets/Editor/Excels";
        
        public static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets,
            string[] movedFromAssetPaths) {
            DataLoad("Test.xlsx");
            
        }

        private static void DataLoad(string fileName) {
            var filePath = $"{tablePath}/{fileName}";
            var fileNameWithoutExtension = fileName.Replace(Path.GetExtension(fileName), "");

            using(FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                IWorkbook workbook = null;

                if(Path.GetExtension(filePath).Equals(".xls")) {
                    workbook = new HSSFWorkbook(stream);
                }
                else {
                    workbook = new XSSFWorkbook(stream);
                }

                // TODO : Change to multi-sheet processable
                var sheet = workbook.GetSheetAt(0);

                IRow settingInfoRow = sheet.GetRow(0);

                var propertyNames = new List<string>();
                
                for(int i = 0; i < settingInfoRow.Cells.Count; i++) {
                    ICell currentCell = settingInfoRow.GetCell(i);
                    propertyNames.Add(currentCell.StringCellValue);
                }

                var rowPairList = new List<HashSet<KeyValuePair<string, string>>>();
                
                for(int i = 1; i <= sheet.LastRowNum; i++) {
                    IRow currentRow = sheet.GetRow(i);

                    var cellPairList = new HashSet<KeyValuePair<string, string>>();
                    
                    for(int j = 0; j < currentRow.Cells.Count; j++) {
                        ICell currentCell = currentRow.GetCell(j);
                        
                        var cellValue = currentCell.CellType switch {
                            CellType.Blank => string.Empty,
                            CellType.Boolean => currentCell.BooleanCellValue.ToString(),
                            CellType.Error => currentCell.ErrorCellValue.ToString(),
                            CellType.Formula => currentCell.CellFormula.ToString(),
                            CellType.Numeric => currentCell.NumericCellValue.ToString(),
                            CellType.String => currentCell.StringCellValue.ToString(),
                            _ => string.Empty,
                        };
                        
                        cellPairList.Add(new KeyValuePair<string, string>(propertyNames[j], cellValue));
                    }
                    
                    rowPairList.Add(cellPairList);
                }
                
                foreach(Assembly assembly in AppDomain.CurrentDomain.GetAssemblies()) {
                    foreach(Type type in assembly.GetTypes()) {
                        var attributes = type.GetCustomAttributes(typeof(DescriptorAttribute), true);

                        if(attributes == null || attributes.Length <= 0) {
                            continue;
                        }

                        var descriptor = attributes[0] as DescriptorAttribute;

                        if(descriptor.TableName.Equals(fileNameWithoutExtension) == false) {
                            continue;
                        }

                        var descriptorType = Type.GetType($"{descriptor.TableName}Descriptor, {assembly.FullName}");

                        if(descriptorType == null) {
                            Debug.Log("Not found Descriptor type.");
                            continue;
                        }

                        var activeDescriptors = new List<object>();
                        
                        for(int i = 0; i < rowPairList.Count; i++) {
                            var descriptorObject = Activator.CreateInstance(descriptorType);

                            foreach(var propertyName in propertyNames) {
                                PropertyInfo propertyInfo = descriptorType.GetProperty(propertyName);

                                var keyValuePair = new KeyValuePair<string, string>();
                                
                                foreach(var pair in rowPairList[i]) {
                                    if(pair.Key != propertyName) {
                                        continue;
                                    }

                                    keyValuePair = pair;
                                    break;
                                }
                                
                                propertyInfo.SetValue(descriptorObject, keyValuePair.Value);
                            }
                            
                            activeDescriptors.Add(descriptorObject);
                        }
                        
                        // TODO : Convert activeDescriptors list to Table Descriptor 파일 
                    }
                }
            }
        }
    }
}