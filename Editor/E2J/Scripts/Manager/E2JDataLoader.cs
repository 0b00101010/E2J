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
    public class E2JDataLoader : AssetPostprocessor{

        private static readonly string tablePath = "Assets/Editor/Excels";
        
        public static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets,
            string[] movedFromAssetPaths) {

            foreach(var importedAsset in importedAssets) {
                if(importedAsset.Contains(tablePath) && (Path.GetExtension(importedAsset).Equals(".xlsx") || Path.GetExtension(importedAsset).Equals(".xls"))) {
                    DataLoad(Path.GetFileName(importedAsset));
                }                
            }
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

                var fieldNames = new List<string>();
                var dataColumns = new List<int>();
                
                for(int i = 0; i < settingInfoRow.Cells.Count; i++) {
                    ICell currentCell = settingInfoRow.GetCell(i);
                    if(currentCell.StringCellValue.StartsWith("_")) {
                        continue;
                    }
                    
                    fieldNames.Add(currentCell.StringCellValue);
                    dataColumns.Add(i);
                }

                var rowPairList = new List<HashSet<KeyValuePair<string, string>>>();
                
                for(int i = 1; i <= sheet.LastRowNum; i++) {
                    IRow currentRow = sheet.GetRow(i);

                    var cellPairList = new HashSet<KeyValuePair<string, string>>();
                    
                    for(int j = 0; j < dataColumns.Count; j++) {
                        ICell currentCell = currentRow.GetCell(dataColumns[j]);

                        if(currentCell == null) {
                            continue;
                        }
                        
                        var cellValue = currentCell.CellType switch {
                            CellType.Blank => string.Empty,
                            CellType.Boolean => currentCell.BooleanCellValue.ToString(),
                            CellType.Error => currentCell.ErrorCellValue.ToString(),
                            CellType.Formula => currentCell.CellFormula.ToString(),
                            CellType.Numeric => currentCell.NumericCellValue.ToString(),
                            CellType.String => currentCell.StringCellValue.ToString(),
                            _ => string.Empty,
                        };
                        
                        cellPairList.Add(new KeyValuePair<string, string>(fieldNames[j], cellValue));
                    }

                    if(cellPairList == null) {
                        continue;
                    }
                    
                    rowPairList.Add(cellPairList);
                }
                
                var activeDescriptors = new List<string>();
                
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

                        for(int i = 0; i < rowPairList.Count; i++) {
                            var descriptorObject = Activator.CreateInstance(descriptorType);

                            foreach(var fieldName in fieldNames) {
                                FieldInfo fieldInfo = descriptorType.GetField(fieldName);

                                if(fieldInfo == null) {
                                    Debug.LogError($"\"{fieldName}\" Field Not found in descriptor.");
                                    return;
                                }

                                var keyValuePair = new KeyValuePair<string, string>();
                                
                                foreach(var pair in rowPairList[i]) {
                                    if(pair.Key != fieldName) {
                                        continue;
                                    }

                                    keyValuePair = pair;
                                    break;
                                }
                                
                                fieldInfo.SetValue(descriptorObject, keyValuePair.Value);
                            }
                            
                            activeDescriptors.Add(JsonUtility.ToJson(descriptorObject));
                        }
                    }
                }
                
                if(activeDescriptors.Count <= 0) {
                    return;
                }
                
                SaveData(fileNameWithoutExtension, activeDescriptors);
            }
        }

        private static void SaveData(string tableName, List<string> jsonList) {
            var directoryPath = new DirectoryInfo( "Assets/Resources/JsonTable");
            
            if(directoryPath.Exists == false) {
                directoryPath.Create();
            }

            var filePath = $"{directoryPath}/{tableName}.txt";
            
            if(File.Exists(filePath)) {
                File.Delete(filePath);
            }

            using(StreamWriter writer = new StreamWriter(filePath)) {
                foreach(var jsonText in jsonList) {
                    writer.WriteLine($"{jsonText}");
                }
            }
        }
    }
}