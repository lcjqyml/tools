package com.milin.tools.maven;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.File;

/**
 * Created by Milin on 2017/5/10.
 */
public class CleanMvnFailedDependencies {

    public static void main(String[] args) {
        String m2Path = "/path/to/your/local/rep";
        if (ArrayUtils.isNotEmpty(args) && StringUtils.isNoneBlank(args[0])) {
            m2Path = args[0];
        }
        findAndDelete(new File(m2Path));
    }

    private static boolean findAndDelete(File file) {
        if (file.exists()) {
            if (file.isFile()) {
                if (file.getName().endsWith("lastUpdated")) {
                    deleteFile(file.getParentFile());
                    return true;
                }
            } else if (file.isDirectory()) {
                File[] files = file.listFiles();
                if (files != null) {
                    for (File f : files) {
                        if (findAndDelete(f)) {
                            break;
                        }
                    }
                }
            }
        }
        return false;
    }

    private static void deleteFile(File file) {
        if (file.exists()) {
            if (file.isFile()) {
                print("删除文件:" + file.getAbsolutePath());
                file.delete();
            } else if (file.isDirectory()) {
                File[] files = file.listFiles();
                if (files != null) {
                    for (File f : files) {
                        deleteFile(f);
                    }
                }
                print("删除文件夹:" + file.getAbsolutePath());
                print("====================================");
                file.delete();
            }
        }
    }

    private static void print(String msg) {
        System.out.println(msg);
    }

}
