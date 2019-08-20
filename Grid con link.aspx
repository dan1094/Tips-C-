<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="Daniel_Delgado_prueba._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
   
<asp:Label ID="OficniasLable" runat="server" Text="Label">Oficinas</asp:Label>
<br />
<br />
    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" >
     <Columns>
        <asp:BoundField DataField="IdOficina" HeaderText="IdOficina" Visible="true" />

        <asp:BoundField HeaderText="NombreOficina" DataField="NombreOficina"></asp:BoundField>

      <asp:TemplateField HeaderText=" ">
        <ItemTemplate>
       <asp:LinkButton runat="server" ID="lnkView" OnClick="lnkView_Click" Text="Ver oficina"></asp:LinkButton>
         </ItemTemplate>
       </asp:TemplateField>


        
        <asp:BoundField HeaderText="Nombre del corresponsal" DataField="NombreCorresponsal"></asp:BoundField>
        <asp:BoundField HeaderText="Numero de giros" DataField="CantidadGiros"></asp:BoundField>

        

    </Columns>

    </asp:GridView>
</asp:Content>
