/*
 * Copyright 2002-2009 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.gageot.excel.core;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.Cell;
import org.gageot.excel.beans.BeanSetter;
import org.gageot.excel.beans.BeanSetterImpl;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.BeanCreationException;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * CallbackHandler implementation that creates a bean of the given class
 * for each row, representing all columns as bean properties.
 *
 * @author David Gageot
 * @see ExcelTemplate#read(String,CellCallbackHandler)
 * @see BeanSetter
 */
public class BeanCellCallbackHandler<T> implements CellCallbackHandler {
	private final Class<T> clazz;
	private final List<T> beans = Lists.newArrayList();
	private final BeanSetter beanSetter;
	private final Map<Integer, String> propertyNames = Maps.newTreeMap();
	private final CellMapper<Object> cellMapper;
	private final CellMapper<String> headerMapper;

	public BeanCellCallbackHandler(Class<T> aClass) {
		this(aClass, new BeanSetterImpl(), new ObjectCellMapper(), new StringCellMapper());
	}

	public BeanCellCallbackHandler(Class<T> clazz,
								   BeanSetter beanSetter,
								   CellMapper<Object> cellMapper, CellMapper<String> headerMapper) {
		this.clazz = clazz;
		this.beanSetter = beanSetter;
		this.cellMapper = cellMapper;
		this.headerMapper = headerMapper;
	}

	public List<T> getBeans() {
		return beans;
	}

	@Override
	public void processCell(Cell cell, int rowNum, int columnNum) throws IOException, BeansException {
		if (0 == rowNum) {
			String propertyName = headerMapper.mapCell(cell, rowNum, columnNum).replaceAll(" ", "_").toLowerCase();
			propertyNames.put(columnNum, propertyName);
			return;
		}

		Object cellValue = cellMapper.mapCell(cell, rowNum, columnNum);
		String propertyName = propertyNames.get(columnNum);

		T bean;
		if (rowNum <= beans.size()) {
			bean = beans.get(rowNum - 1);
		} else {
			bean = createBean(clazz);
			beans.add(bean);
		}

		beanSetter.setProperty(bean, propertyName, cellValue);
	}

	/**
	 * Create a bean for a given class.
	 * Default strategy is to create an empty bean calling its empty
	 * constructor. This method can be overridden for a different strategy.
	 * @param clazz the class of the bean to create
	 */
	protected T createBean(Class<T> aClazz) throws BeansException {
		try {
			return aClazz.newInstance();
		} catch (InstantiationException e) {
			throw new BeanCreationException("Impossible to create bean", e);
		} catch (IllegalAccessException e) {
			throw new BeanCreationException("Impossible to create bean", e);
		}
	}
}
