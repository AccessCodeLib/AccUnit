﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Dieser Code wurde von einem Tool generiert.
//     Laufzeitversion:4.0.30319.42000
//
//     Änderungen an dieser Datei können falsches Verhalten verursachen und gehen verloren, wenn
//     der Code erneut generiert wird.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AccessCodeLib.AccUnit.VbeAddIn.Resources {
    using System;
    
    
    /// <summary>
    ///   Eine stark typisierte Ressourcenklasse zum Suchen von lokalisierten Zeichenfolgen usw.
    /// </summary>
    // Diese Klasse wurde von der StronglyTypedResourceBuilder automatisch generiert
    // -Klasse über ein Tool wie ResGen oder Visual Studio automatisch generiert.
    // Um einen Member hinzuzufügen oder zu entfernen, bearbeiten Sie die .ResX-Datei und führen dann ResGen
    // mit der /str-Option erneut aus, oder Sie erstellen Ihr VS-Projekt neu.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class MessageStrings {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal MessageStrings() {
        }
        
        /// <summary>
        ///   Gibt die zwischengespeicherte ResourceManager-Instanz zurück, die von dieser Klasse verwendet wird.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("AccessCodeLib.AccUnit.VbeAddIn.Resources.MessageStrings", typeof(MessageStrings).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Überschreibt die CurrentUICulture-Eigenschaft des aktuellen Threads für alle
        ///   Ressourcenzuordnungen, die diese stark typisierte Ressourcenklasse verwenden.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Sucht eine lokalisierte Zeichenfolge, die The application is not saved in a fully compiled state. ähnelt.
        /// </summary>
        internal static string Application_not_saved_in_compiled_state {
            get {
                return ResourceManager.GetString("Application_not_saved_in_compiled_state", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Sucht eine lokalisierte Zeichenfolge, die Unknown application reference.
        ///Please call &apos;Application.VBE.AddIns(&quot;AccUnit.GUI.Connect&quot;).Object.Application = Application&apos; ähnelt.
        /// </summary>
        internal static string MissingHostApplicationReference {
            get {
                return ResourceManager.GetString("MissingHostApplicationReference", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Sucht eine lokalisierte Zeichenfolge, die &apos;{0}&apos; no longer exists.
        ///Please update the test list. ähnelt.
        /// </summary>
        internal static string MissingTestClassInVBProject {
            get {
                return ResourceManager.GetString("MissingTestClassInVBProject", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Sucht eine lokalisierte Zeichenfolge, die Tests exported:
        ///{0} ähnelt.
        /// </summary>
        internal static string TestExportedCommitMessage {
            get {
                return ResourceManager.GetString("TestExportedCommitMessage", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Sucht eine lokalisierte Zeichenfolge, die Tests imported ähnelt.
        /// </summary>
        internal static string TestImportedCommitMessage {
            get {
                return ResourceManager.GetString("TestImportedCommitMessage", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Sucht eine lokalisierte Zeichenfolge, die The content of the previous message has been copied to the clipboard. ähnelt.
        /// </summary>
        internal static string UITools_ShowException_message_has_been_copied_to_the_clipboard {
            get {
                return ResourceManager.GetString("UITools_ShowException_message_has_been_copied_to_the_clipboard", resourceCulture);
            }
        }
    }
}
