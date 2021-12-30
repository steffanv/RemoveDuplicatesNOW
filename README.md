# RemoveDuplicatesNOW
Remove Duplicates in your Outlook Calendar (NOW!).

You imported holidays or birthdays by accident twice? No easy solution is available?

Before:
![before](https://user-images.githubusercontent.com/7317719/147685733-b8cca266-65dd-4b0c-980e-b50d255ff308.jpg)

After:
![after](https://user-images.githubusercontent.com/7317719/147685757-498b7db0-90a2-4c13-8a57-eb5ef47baaf2.jpg)

What is my definition of a duplicate?

```c#
public bool Equals(AppointmentItem x, AppointmentItem y)
{
    if (x.Subject == y.Subject && x.Start == y.Start && x.End == y.End && x.Body == y.Body)
    {
        return true;
    }
    else
    {
        return false;
    }
}
```
Please refer to the Add-In ribbon.
![outlook](https://user-images.githubusercontent.com/7317719/147686266-6e731adc-92c7-450b-98ce-e9ffa2f00978.jpg)

The removed duplicates can be found in the Deleted Items folder.

The software is provided "as is" and you use it on your own risk. Please refer to the Apache License.


