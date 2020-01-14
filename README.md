<h2>Metadata PCF library</h3>

<h3>Summary</h3>
This is a set of PCF (Power Component Framework) controls to provide user friendly access to Powerplatform metadata allowing you to easily select entities, attributes, lookup values and views and store them in standard PowerApp text fields for use elsewhere.

There were built to improve and remove the complex javascript that we used to need on the administration forms that we use across a number of our solutions and are shared here to ensure other people do not need to waste time rewriting similar logic.

If you do not wish to compile the controls manually you can download an [unmanaged solution](./MetaDataSelectPCF_1_0_0_1.zip)

<h3> The controls</h3>
The solution contains the following controls:-


**Select Entity**

A drop down / select list that allows you to select an entity and store it's logical name. Optionally it you can bind a second text field to store the entity's display name.  

2 Properties:-  
Entity - *Required* Bound singleline text field  
Display Name - *Optional* Bound singleline text field.

![Entity Select Demo!](Entity%20Select.gif)

**Select Attribute**

A drop down / select list that displays  and allows you to select an attribute from the specified entity. The control stores the logical name of the field / attribute and displays it in a user friendly format *Display Name (logical name)*.

2 properties:-  
Attribute - *Required* Bound singleline text field  
EntityName - Singleline text field input (usually set to the Select Entity field elsewhere on the form)

![Attrbute Select Demo!](Lookup%20Select.gif)

**Select Lookup**

A variation of the Select Attribute control above that only shows and allows you to select lookup fields from the specified entity. The control stores the logical name of the lookup field and  displays it in a user friendly format - *Display Name (logical name)*  
2 properties:-  
Attribute - *Required* Bound singleline text field  
EntityName - Singleline text field input (usually set to the Select Entity field elsewhere on the form)

**Select View**

A drop down / select list that displays and allows you to select a view  from the specified entity. The control stores the Guid value of the View and displays the views user friendly name.
 
2 properties:-  
Attribute - *Required* Bound singleline text field  
EntityName - Singleline text field input (usually set to the Select Entity field elsewhere on the form)

**Multiple Attributes**

A drop down / select list that provides a multiselect combo box where you can select multiple attributes from the specified entity. The control stores the logical name of the selected attributes in a comma seperated list and displays them in a user friendly format *Display Name (logical name)*.

2 properties:-  
Attribute - *Required* Bound singleline text field  
EntityName - Singleline text field input (usually set to the Select Entity field elsewhere on the form)

![Multiple Attribute Select Demo!](multiselect.gif)

**Select Related**

A variation of the Select Lookup drop down / select list used to select the preferred join between two entities (Form entity and Query entity). The dropdown list shows the joins to the Form entity on the Query entity.
<ul>
<li>Where no relationship exists, an error message is shown.  
<li>Where a single relationship exists, the control will automatically select it.  
<li>Where multiple options exist the value will be blank until one is selected.  
</ul>
The control stores the logical name of the lookup field and  displays it in a user friendly format - *Display Name (logical name)*  

3 properties:-  
Attribute - *Required* Bound singleline text field  
FormEntity -  *Required* Singleline Bound text field 
QueryEntity -  *Required* Singleline Bound text field 


<h3>Notes</h3>

The controls support readonly / disabled mode but do not support Field level security as I cannot think of any use case that would require it.  

CSS styling is a variation of Benedikt Bergmann's work see https://benediktbergmann.eu/2020/01/04/pcf-css-for-model-driven-apps-mda/  
with additional Javascript to more accurately follow the way right hand arrows work within Microsoft's standard controls see https://powerusers.microsoft.com/t5/Power-Apps-Component-Framework/PCF-Control-hasFocus/m-p/442274#M1544 .  

Error handling uses setNotification and clearNotification available within the Utlity feature of PCF framework  
~~~
(<any>this.context.utils).setNotification("No relationship between selected Query ("+ this.queryEntity +") and Form ("+ this.formEntity +") entities.","123")    
~~~
and
~~~
     (<any>this.context.utils).clearNotification("123");    
~~~
to clear the error

To access the error handling just enable the following in the ControlManifest file 

~~~
   <feature-usage>    
      <uses-feature name="Utility" required="true" />
    </feature-usage>
~~~

