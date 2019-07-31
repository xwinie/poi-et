package com.jg.poiet.render;

import com.jg.poiet.config.ELMode;

public class RenderFactory {

    public static Render getRender(Object model, ELMode mode) {
        Render render;
        switch (mode) {
            case POI_TL_STICT_MODE:
                render = new Render(new ELObjectRenderDataCompute(model, true));
                break;
            default:
                render = new Render(model);
                break;
        }
        return render;
    }

}
