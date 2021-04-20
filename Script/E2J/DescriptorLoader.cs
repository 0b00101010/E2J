using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace E2J {
    public class DescriptorLoader : MonoBehaviour {
        private static DescriptorLoader instance;
        public static DescriptorLoader Instance {
            get {
                if(instance == null) {
                    var obj = GameObject.FindObjectOfType<DescriptorLoader>(true);

                    if(obj == null) {
                        var prefab = Resources.Load<DescriptorLoader>($"Prefabs/DescriptorLoader");

                        if(prefab != null) {
                            obj = prefab;
                            Instantiate(prefab.gameObject, Vector2.zero, Quaternion.identity);
                        }
                        else {
                            obj = new GameObject("DescriptorLoader").AddComponent<DescriptorLoader>();

                            var directoryPath = "Assets/Resources/Prefabs";
                            if(Directory.Exists(directoryPath) == false) {
                                Directory.CreateDirectory(directoryPath);
                            }

                            var prefabPath = $"{directoryPath}/DescriptorLoader.prefab";
                            prefabPath = AssetDatabase.GenerateUniqueAssetPath(prefabPath);

                            PrefabUtility.SaveAsPrefabAssetAndConnect(obj.gameObject, prefabPath, InteractionMode.UserAction);
                        }
                    }

                    instance = obj;
                }

                return instance;
            }
        }
        
        public T[] ConvertDescriptor<T>() {
            var attribute = (DescriptorAttribute)Attribute.GetCustomAttribute(typeof(T), typeof(DescriptorAttribute));

            if(attribute is null) {
                Debug.Log($"Descriptor Attribute Not Found: {nameof(T)}");
                return null;
            }
            
            var jsonTextFile = Resources.Load<TextAsset>($"JsonTable/{attribute.TableName}");

            if(jsonTextFile == null) {
                Debug.Log($"Not Found Json Table File: {nameof(T)}");
                return null;
            }
            
            string[] jsonStrings = jsonTextFile.text.Split('\n');
            
            var descriptors = new List<T>();            
            
            foreach(var jsonString in jsonStrings) {
                if(jsonString.Equals(string.Empty)) {
                    continue;
                }
                
                T descriptor = JsonUtility.FromJson<T>(jsonString);
                descriptors.Add(descriptor);
            }
            
            return descriptors.ToArray();
        }
    }
}