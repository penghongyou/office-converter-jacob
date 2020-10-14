package cn.com.taiji.converter.vo;

import com.jacob.activeX.ActiveXComponent;
import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain = true)
public class ActiveXComVo {
    private ActiveXComponent app;
    private String wordPid;
    private Long lastUpdateTime;
    private String currentThreadName;
}
