/**
 * 主题统一入口
 */

export type { Theme, ThemeColors, FontConfig, PageLayout, ParagraphFormat, HeadingStyle, TableStyle, WordStyles, ExcelStyles } from './types.js'
export { cmToTwips, createTheme } from './base.js'

import { governmentTheme } from './government.js'
import { academicTheme } from './academic.js'
import { softwareDocTheme } from './software.js'
import { businessTheme } from './business.js'
import { alibabaTheme } from './alibaba.js'
import { tencentTheme } from './tencent.js'
import { bytedanceTheme } from './bytedance.js'
import { minimalTheme } from './minimal.js'
import { Theme } from './types.js'

export {
  governmentTheme, academicTheme, softwareDocTheme, businessTheme,
  alibabaTheme, tencentTheme, bytedanceTheme, minimalTheme
}

const defaultTheme = businessTheme

const themes: Record<string, Theme> = {
  government: governmentTheme,
  academic: academicTheme,
  software: softwareDocTheme,
  business: businessTheme,
  alibaba: alibabaTheme,
  tencent: tencentTheme,
  bytedance: bytedanceTheme,
  minimal: minimalTheme,
  default: defaultTheme
}

export function getTheme(name?: string): Theme {
  return themes[name || 'default'] || defaultTheme
}

export function listThemes(): Array<{ name: string; displayName: string; description: string }> {
  return Object.values(themes)
    .filter(t => t.name !== 'default')
    .map(t => ({ name: t.name, displayName: t.displayName, description: t.description }))
}
