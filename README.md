## NAME

genomap2ldap.pl - loads information about GenoMap individuals into LDAP directory

## SYNOPSIS

    genomap2ldap.pl -f genomap.gno [-h ldap.host.com] [-D bind_dn] [-w bind_password] [-S users_dn] [-G groups_dn]
    genomap2ldap.pl --help

    genomap2ldap.pl -f C:/Documents/myfile.gno -S cn=user,cn=domain -G cn=groups,cn=domain -h 192.168.1.5 -D cn=ldapadmin,cn=domain -w superUserPass

## OPTIONS

- __\--help__

    Prints this help message.

- __\--host|-h__

    Optionally specify the LDAP host to connect to. If not defined, the data is printed to standard output in LDIFF format,
    which can be manually imported into LDAP directory.

- __\--search-dn|-S__

    Optionally specify the search DN, which should be used during the lookups for individuals. If this is not defined,
    all lookups are made in root DN ("").

- __\--search-group-dn|-G__

    Optionally specify the search DN, which should be used during the lookups for groups. If this is not defined,
    all lookups are made in root DN ("").

- __\--bind-dn|-D__

    Optionally specify the bind DN, which should be used during the authentication phase. Note, that both bind DN and bind password
    should be specified, otherwise the anonymous bind will take place.

- __\--bind-pass|-w__

    Optionally specify the bind password, which should be used during the authentication phase.

- __\--no-load-images__

    Do not load the images.

## DESCRIPTION

The GenoMap XML data is parsed and converted to .ldiff format. The program prints the resulting .ldiff to the screen, if
no LDAP host is provided. Otherwise for each added individual it is checked, if the entry already exists in LDAP
directory. If positive, the entry fields are updated, otherwise a new entry is added.

## LDAP Attributes Used by Thunderbird

> birthday o company mail modifytimestamp mozillaUseHtmlMail xmozillausehtmlmail mozillaCustom2 custom2
> mozillaHomeCountryName ou department departmentnumber orgunit mobile cellphone carphone telephoneNumber title
> mozillaCustom1 custom1 sn surname mozillaNickname xmozillanickname mozillaWorkUrl workurl labeledURI
> facsimiletelephonenumber fax mozillaSecondEmail xmozillasecondemail mozillaCustom4 custom4 nsAIMid nscpaimscreenname
> street streetaddress postOfficeBox givenName l locality homePhone mozillaHomeUrl homeurl mozillaHomeStreet st region
> mozillaHomePostalCode mozillaHomeLocalityName mozillaCustom3 custom3 birthyear mozillaWorkStreet2 mozillaHomeStreet2
> postalCode zip birthmonth c countryname pager pagerphone mozillaHomeState description notes cn commonname objectClass

More information:

* http://wiki.mozilla.org/MailNews:Mozilla%20LDAP%20Address%20Book%20Schema
* http://www.mozilla.org/projects/thunderbird/specs/ldap.html
* http://www.genopro.com/
