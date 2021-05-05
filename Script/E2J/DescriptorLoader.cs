using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace E2J {
    public static class DescriptorLoader {
        public static T[] ConvertDescriptor<T>() {
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
            
            var descriptors = new T[jsonStrings.Length];

            for(int i = 0; i < jsonStrings.Length; i++) {
                if(jsonStrings[i].Equals(string.Empty)) {
                    continue;
                }

                T descriptor = JsonUtility.FromJson<T>(jsonStrings[i]);
                descriptors[i] = descriptor;
            }

            return descriptors;
        }
    }
}