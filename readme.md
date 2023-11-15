  
<span style="font-family:Arial sans-serif;font-size:16px;">WD Pilotage de Excel</span>

  
<span style="font-family:Arial sans-serif;font-size:14px;">Cet exemple montre comment piloter Excel via OLE Automation. </span>

<span style="font-family:Arial sans-serif;font-size:14px;">Toutes les fonctions d'Excel peuvent être ainsi pilotées.</span>

  
<span style="font-family:Arial sans-serif;font-size:14px;">Cet exemple nécessite une version Excel 97 ou supérieure.</span>

  
<span style="font-family:Arial sans-serif;font-size:14px;">Résumé de l'exemple livré avec WINDEV : </span>

<span style="font-family:Arial sans-serif;font-size:14px;">Piloter un logiciel bureautique comme Excel peut s'avérer utile pour analyser des données provenant d'une base de données. </span>

<span style="font-family:Arial sans-serif;font-size:14px;">Grâce à la classe "CExcel" livrée avec WINDEV, ce traitement devient très simple.</span>

<span style="font-family:Arial sans-serif;font-size:14px;">Les principales fonctions de Excel sont directement pilotables (graphe, insertion d'objets, tris...).</span>

<span style="font-family:Arial sans-serif;font-size:14px;">Comment piloter Excel via OLE Automation ?</span>

<span style="font-family:Arial sans-serif;font-size:14px;">Un objet OLE Automation dispose de méthodes et de propriétés. Ceci permet de le piloter directement en WLangage.</span>

<span style="font-family:Arial sans-serif;font-size:14px;">Par exemple, pour mettre la cellule sélectionnée en gras :</span>

<span style="font-family:Arial sans-serif;font-size:14px;">MonObjetOLEAutomation&gt;&gt;Selection&gt;&gt;Font&gt;&gt;Bold = Vrai </span>

  
  
<span style="font-family:Arial sans-serif;font-size:14px;">( \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_ ° \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_ )</span>

  
<span style="font-family:Arial sans-serif;font-size:10px;">Conditions d'utilisation :</span>

<span style="font-family:Arial sans-serif;font-size:10px;">Le programme est fourni dans un but didactique.</span>

<span style="font-family:Arial sans-serif;font-size:10px;">L'utilisation de ce programme s'effectue sous la responsabilité de son utilisateur. La responsabilité de PC SOFT ne pourra en aucun cas être mise en cause si le programme ne fonctionne pas tel que vous l'attendez, ou pour quelque raison que ce soit. </span>

<span style="font-family:Arial sans-serif;font-size:10px;">Tout détenteur d'une licence WINDEV 28 enregistrée est autorisé à modifier ce programme, et est autorisé à l'inclure dans une ou plusieurs de ses applications. Dans ces cas, toute mention se rapportant à PC SOFT, à WinDev ou à WebDev devra être supprimée, afin qu'aucun doute ne puisse subsister dans l'esprit d'un utilisateur final.</span>

<span style="font-family:Arial sans-serif;font-size:10px;">Il est interdit de diffuser ou vendre ce programme exemple sans modification substantielle, ou sans l'inclure dans une application de taille substantielle.</span>

<span style="font-family:Arial sans-serif;font-size:10px;">Le support technique pour ce programme exemple est accessible à travers le service "Assistance Directe" uniquement.</span>

<span style="font-family:Arial sans-serif;font-size:10px;">Attention: si plusieurs personnes sont susceptibles d'avoir consulté ce programme, il se peut que le programme ait été modifié! </span>

<span style="font-family:Arial sans-serif;font-size:10px;">Assurez-vous d'étudier une version originale de ce programme.</span>

  
  