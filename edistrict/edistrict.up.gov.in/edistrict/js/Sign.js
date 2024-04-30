// JScript File

// CAPICOM constants

var CAPICOM_STORE_OPEN_READ_ONLY = 0;

var CAPICOM_CURRENT_USER_STORE = 2;

var CAPICOM_CERTIFICATE_FIND_SHA1_HASH = 0;

var CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY = 6;

var CAPICOM_CERTIFICATE_FIND_TIME_VALID = 9;

var CAPICOM_CERTIFICATE_FIND_KEY_USAGE = 12;

var CAPICOM_DIGITAL_SIGNATURE_KEY_USAGE = 0x00000080;

var CAPICOM_AUTHENTICATED_ATTRIBUTE_SIGNING_TIME = 0;

var CAPICOM_INFO_SUBJECT_SIMPLE_NAME = 0;

var CAPICOM_ENCODE_BASE64 = 0;

var CAPICOM_ENCODE_BINARY = 1;

var CAPICOM_E_CANCELLED = -2138568446;

var CERT_KEY_SPEC_PROP_ID = 6;

var CAPICOM_CERT_INFO_TYPE = 6;

var CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME = 4;

function IsCAPICOMInstalled()

{

    if (typeof(oCAPICOM) == "object")

    {

        if ((oCAPICOM.object != null))

        {

            // We found CAPICOM!

            return true;

        }

    }

}



function init()

{

    // Filter the certificates to only those that are good for my purpose

    var FilteredCertificates = FilterCertificates();



    // if only one certificate was found then lets show that as the selected certificate

    if (FilteredCertificates.Count == 1)

    {

        txtCertificate.value = FilteredCertificates.Item(1).GetInfo(CAPICOM_INFO_SUBJECT_SIMPLE_NAME);

        txtCertificate.hash = FilteredCertificates.Item(1).Thumbprint;

    } else

    {

        txtCertificate.value = "";

        txtCertificate.hash = "";

    }



    // clean up

    FilteredCertificates = null;

}





function FilterCertificates()

{

    // instantiate the CAPICOM objects

    var MyStore = new ActiveXObject("CAPICOM.Store");

    var FilteredCertificates = new ActiveXObject("CAPICOM.Certificates");



    // open the current users personal certificate store

    try

    {

        MyStore.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY);

    } catch (e)

    {

        if (e.number != CAPICOM_E_CANCELLED) {
            alert("An error occurred while opening your personal certificate store, aborting");

            return false;
        }



    }



    // find all of the certificates that:

    //   * Are good for signing data

    //       * Have PrivateKeys associated with then - Note how this is being done :)

    //   * Are they time valid

    var FilteredCertificates = MyStore.Certificates.Find(CAPICOM_CERTIFICATE_FIND_KEY_USAGE, CAPICOM_DIGITAL_SIGNATURE_KEY_USAGE).Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID).Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CERT_KEY_SPEC_PROP_ID);

    return FilteredCertificates;



    // Clean Up

    MyStore = null;

    FilteredCertificates = null;

}



function FindCertificateByHash(szThumbprint)

{

    // instantiate the CAPICOM objects

    var MyStore = new ActiveXObject("CAPICOM.Store");



    // open the current users personal certificate store

    try

    {

        MyStore.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY);

    } catch (e)

    {

        if (e.number != CAPICOM_E_CANCELLED)

        {

            alert("An error occurred while opening your personal certificate store, aborting");

            return false;

        }

    }



    // find all of the certificates that have the specified hash

    var FilteredCertificates = MyStore.Certificates.Find(CAPICOM_CERTIFICATE_FIND_SHA1_HASH, szThumbprint);

    return FilteredCertificates.Item(1);



    // Clean Up

    MyStore = null;

    FilteredCertificates = null;

}



function btnSelectCertificate_OnClick()

{

    var mn = document.getElementById("txtCertificate");
    var cstinfo = document.getElementById("cstinfo");
    var currdate = document.getElementById("currdate");

    // retrieve the filtered list of certificates

    var FilteredCertificates = FilterCertificates();



    // if only one certificate was found then lets show that as the selected certificate

    if (FilteredCertificates.Count >= 1)

    {
        try

        {
            // Pop up the selection UI
            var SelectedCertificate = FilteredCertificates.Select();

            var a = new Date(new Date(SelectedCertificate.Item(1).ValidToDate));
            var todayDate = new Date(currdate.value);



            //                            if (a >= todayDate) {
            //                                alert('hello');
            //                            }
            //                            else {
            //                                alert("Invalid certificate or certificate expired ");
            //                                return;
            //                            }


            if (a < todayDate) {

                alert("Digital Signature certificate expired ");
                SelectedCertificate = '';
                return;
            }


            mn.value = SelectedCertificate.Item(1).GetInfo(CAPICOM_INFO_SUBJECT_SIMPLE_NAME);
            mn.hash = SelectedCertificate.Item(1).Thumbprint;
            cstinfo.value = SelectedCertificate.Item(1).SubjectName;

            //alert('1' +SelectedCertificate.Item(1).ValidToDate);







            //alert('<%=date%>');
            //NotAfter
            //  txtCertificate.value = SelectedCertificate.Item(1).GetInfo(CAPICOM_INFO_SUBJECT_SIMPLE_NAME);

            //  txtCertificate.hash = SelectedCertificate.Item(1).Thumbprint;
        } catch (e)

        {
            alert(e.description);
            txtCertificate.value = "";

            txtCertificate.hash = "";

        }

    } else

    {

        alert("You have no valid certificates to select from");

    }



    // Clean-Up

    SelectedCertificate = null;

    FilteredCertificates = null;

}



function btnVerifySig_OnClick()

{

    var txtsigndata = document.getElementById("txtsigndata");
    var txtPlainText = document.getElementById("txtPlainText");


    // instantiate the CAPICOM objects
    var SignedData = new ActiveXObject("CAPICOM.SignedData");

    var signature = txtsigndata.value;

    var content = txtPlainText.value;


    SignedData.Content = content;


    try

    {

        var szverify = SignedData.Verify(signature, true);

        alert("Signature Verified");

    } catch (e)

    {

        if (e.number != CAPICOM_E_CANCELLED)

        {

            alert("An error occurred when attempting to verify the content, the errot was: " + e.description);

            return false;

        }

    }



}

function btnSignedData_OnClick()

{

    // instantiate the CAPICOM objects



    var SignedData = new ActiveXObject("CAPICOM.SignedData");
    var Signer = new ActiveXObject("CAPICOM.Signer");
    var TimeAttribute = new ActiveXObject("CAPICOM.Attribute");
    var txtCertificate = document.getElementById("txtCertificate");
    var txtsigndata = document.getElementById("txtsigndata");
    var txtPlainText = document.getElementById("txtPlainText");

    // CAPICOM.Certificate.ICertificate2
    var certd = new ActiveXObject("CAPICOM.Certificate");
    // 
    //          
    // only do this if the user selected a certificate

    if (txtCertificate.hash != "")

    {


        // Set the data that we want to sign


        try

        {
            SignedData.Content = txtPlainText.value;
            //System.Enum.GetName

            //var info= certd.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME);
            //var info= certd.IssuerName;

            //alert(info);

            // Set the Certificate we would like to sign with
            Signer.Certificate = FindCertificateByHash(txtCertificate.hash);

            // Set the time in which we are applying the signature

            var Today = new Date();

            TimeAttribute.Name = CAPICOM_AUTHENTICATED_ATTRIBUTE_SIGNING_TIME;
            TimeAttribute.Value = Today.getVarDate();
            Today = null;
            Signer.AuthenticatedAttributes.Add(TimeAttribute);


            // Do the Sign operation

            var szSignature = SignedData.Sign(Signer, true, CAPICOM_ENCODE_BASE64);


        } catch (e)

        {

            if (e.number != CAPICOM_E_CANCELLED)

            {

                alert("An errork occurred when attempting to sign the content, the errot was: " + e.description);

                return false;

            }




        }



        txtsigndata.value = szSignature;
        var h1 = document.getElementById("HF1");
        h1.value = szSignature;
        var tabname = document.getElementById("tab2");
        tabname.style.display = 'none';
        var digitalID = document.getElementById("digitalID");
        digitalID.value = txtCertificate.value
    } else

    {

        alert("No Certificate has been selected.");

    }

}