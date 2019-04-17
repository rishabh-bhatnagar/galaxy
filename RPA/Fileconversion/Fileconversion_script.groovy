import com.workfusion.studio.rpa.recorder.api.*
import com.workfusion.studio.rpa.recorder.api.types.*
import com.workfusion.studio.rpa.recorder.api.custom.*
import com.workfusion.studio.rpa.recorder.api.internal.representation.*

import com.workfusion.bot.exception.*

        
            def image_found_edit = RBoolean.fromCanonical('false')
            
        
            def doc_files = RList.of()
            

com.workfusion.rpa.helpers.RPA.metaClass.static.$ = { Closure c -> c.call() } // Support for Expression action. Should be implemented in RPA class in next release.

enableTypeOnScreen()

    def i0 = com.workfusion.rpa.helpers.resources.Filter
        .files()        
        .get()
        
    doc_files = RList.of(Resource.listFolder("C:\\Users\\Shruti\\Desktop\\Docs", i0))


    doc_files.each( {element ->
        


inDesktop {
        sendKeys(StringTransformations.getHotKeyText(114, 4))
}

        


inDesktop {
            
            sendKeys(StringTransformations.escapeAutoitText(element.toString()))
}

        


inDesktop {
        sendKeys(StringTransformations.getKeyPressText(28, 13, 10, 0))
}

        
    
    def i2 = System.currentTimeMillis() + 5000
    
    boolean i1 = true

    while (i1) {
        try {
            $(byImage("1545798048865-anchor.apng")).getCoordinates()
            image_found_edit = RBoolean.TRUE
            break
        } catch (Exception ignored) {
            image_found_edit = RBoolean.FALSE
        }
        i1 = System.currentTimeMillis() < i2 
    } 
    

        if ((image_found_edit) == (RBoolean.fromCanonical('true'))) {
     

                $(byImage("1545798148435-anchor-1545798148440.apng", Integer.valueOf(0), Integer.valueOf(0))).click()

} else {
     
}
        
    
    def i4 = System.currentTimeMillis() + 200
    
    boolean i3 = true

    while (i3) {
        try {
            $(byImage("1545794609167-anchor.apng")).getCoordinates()
            
            break
        } catch (Exception ignored) {
            
        }
        i3 = System.currentTimeMillis() < i4 
    } 
    

        if ((image_found_edit) == (RBoolean.fromCanonical('true'))) {
     

                $(byImage("1545794609092-anchor-1545794681244.apng", Integer.valueOf(0), Integer.valueOf(0))).click()

} else {
     
}
        


inDesktop {
        sendKeys(StringTransformations.getKeyPressText(88, 0, 123, 0))
}

        
    
    def i6 = System.currentTimeMillis() + 100
    
    boolean i5 = true

    while (i5) {
        try {
            $(byImage("1545795454385-anchor.apng")).getCoordinates()
            image_found_edit = RBoolean.TRUE
            break
        } catch (Exception ignored) {
            image_found_edit = RBoolean.FALSE
        }
        i5 = System.currentTimeMillis() < i6 
    } 
    

        if ((image_found_edit) == (RBoolean.fromCanonical('true'))) {
     

                $(byImage("1545795454380-anchor-1545795612470.apng", Integer.valueOf(0), Integer.valueOf(0))).click()

} else {
     

                $(byImage("1545794954758-anchor-1545794954766.apng", Integer.valueOf(0), Integer.valueOf(0))).click()

}
        

                $(byImage("1545794981846-anchor-1545794981854.apng", Integer.valueOf(0), Integer.valueOf(0))).click()

        


inDesktop {
        sendKeys(StringTransformations.getKeyPressText(28, 13, 10, 0))
}

            sleep(2000)

        

                $(byImage("1545795322700-anchor-1545795322706.apng", Integer.valueOf(0), Integer.valueOf(0))).click()

    })

